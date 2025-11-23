namespace GenerateAzureStencilsFromSvgIcons.Object;

using System.Globalization;
using GenerateAzureStencilsFromSvgIcons.Model;
using Visio = Microsoft.Office.Interop.Visio;

public sealed class VisioStencilGenerator
{
    private readonly AppPathsConfigModel _paths;
    private readonly NameFormatter _nameFormatter;
    private static string? _fontType;
    private static int _fontSize;

    public VisioStencilGenerator(AppPathsConfigModel paths, NameFormatter nameFormatter, string fontType, int fontSize)
    {
        _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        _nameFormatter = nameFormatter ?? throw new ArgumentNullException(nameof(nameFormatter));
        _fontType = fontType;
        _fontSize = fontSize;
    }

    public void GenerateStencils(IEnumerable<IGrouping<string?, string>> groups)
    {
        Visio.Application? visioApp = null;

        string dateFolder = DateTime.Today.ToString("yyyy-MM-dd");
        var finalPath = Path.Combine(_paths.OutputFolder, dateFolder);
        FileHelper.DeleteDirectoryIfExists(finalPath);
        Directory.CreateDirectory(finalPath);

        try
        {
            visioApp = new Visio.Application
            {
                Visible = false
            };

            var groupList = groups.ToList();
            Console.WriteLine($"Category folders (stencils) to be created: {groupList.Count}");
            Console.WriteLine();

            int groupIndex = 0;

            foreach (var group in groupList)
            {
                groupIndex++;

                string folderPath = group.Key ?? _paths.IconsRoot;
                string rawCategory = Path.GetFileName(folderPath);
                if (string.IsNullOrWhiteSpace(rawCategory))
                    rawCategory = "(Root)";

                string prettyCategory = _nameFormatter.ToFolderTitleCase(rawCategory);
                string stencilName = $"Azure-{prettyCategory}.vssx";
                string stencilPath = Path.Combine(finalPath, stencilName);

                Console.WriteLine($"[{groupIndex} / {groupList.Count}] Folder     : '{rawCategory}'");
                Console.WriteLine($"[{groupIndex} / {groupList.Count}] Stencil    : '{stencilName}'");

                var svgFiles = group.ToList();
                if (!svgFiles.Any())
                {
                    Console.WriteLine("  (No SVGs in this folder, skipping.)");
                    Console.WriteLine();
                    continue;
                }

                var stencilDoc = visioApp.Documents.AddEx(
                    string.Empty,
                    Visio.VisMeasurementSystem.visMSDefault,
                    (short)Visio.VisOpenSaveArgs.visAddStencil,
                    0);

                var masters = stencilDoc.Masters;

                int i = 0;
                foreach (var svgPath in svgFiles)
                {
                    i++;
                    string fileName = Path.GetFileName(svgPath) ?? svgPath;
                    Console.WriteLine($"  [{i} / {svgFiles.Count}] Importing {fileName}...");

                    string masterName = _nameFormatter.GetMasterNameFromFile(fileName);

                    Visio.Master master = masters.Add();
                    master.Name = masterName;

                    Visio.Shape importedShape = master.Import(svgPath);

                    double oldWidthIu = importedShape.CellsU["Width"].ResultIU;
                    importedShape.CellsU["Width"].FormulaU = "0.7 in";
                    importedShape.CellsU["Height"].FormulaU =
                        (importedShape.CellsU["Height"].ResultIU * (0.7 / oldWidthIu))
                        .ToString(CultureInfo.InvariantCulture) + " in";

                    Visio.Shape targetShape = FindTopLevelGroup(master) ?? importedShape;

                    InitializeMasterShape(targetShape, masterName);
                }

                stencilDoc.SaveAs(stencilPath);
                stencilDoc.Close();

                Console.WriteLine($"  Completed stencil: {stencilPath}");
                Console.WriteLine();
            }

            Console.WriteLine("All stencils created.");
            Console.WriteLine($"Open Visio and check: More Shapes -> My Shapes -> " +
                              $"{Path.GetFileName(_paths.OutputFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))}" +
                              $" -> {dateFolder} -> Azure-*.vssx");
            Console.WriteLine($"Azure SVG Icons directory: `{_paths.IconsRoot}`");
            Console.WriteLine($"Azure SVG Stencils directory: `{finalPath}`");
            Console.WriteLine($"By using these stencils/icons from Microsoft, you agree to their terms of use: `{_paths.IconsRoot}\\Azure_Public_Service_Icons\\{dateFolder}\\Microsoft_Terms_of_Use.pdf`");
        }
        finally
        {
            VisioHelper.SafeQuit(visioApp);
        }
    }

    private static Visio.Shape? FindTopLevelGroup(Visio.Master master)
    {
        foreach (Visio.Shape s in master.Shapes)
        {
            if (s.Type == (short)Visio.VisShapeTypes.visTypeGroup)
                return s;
        }

        return null;
    }

    private static void InitializeMasterShape(Visio.Shape shape, string masterName)
    {
        var main = shape;
        SetGroupSelectionMode(main);
        ClearDoubleClickEvents(main);
        var label = DetermineLabelShape(main);
        if (label != null)
            ConfigureLabelShape(label, masterName);
        AddConnectionPoints(main);
    }

    private static void SetGroupSelectionMode(Visio.Shape main)
    {
        try
        {
            if (main.Type == (short)Visio.VisShapeTypes.visTypeGroup)
            {
                var cell = main.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowGroup,
                    (short)Visio.VisCellIndices.visGroupSelectMode];
                cell.FormulaU = "0"; // group-only
            }
        }
        catch
        {
        }
    }

    private static void ClearDoubleClickEvents(Visio.Shape main)
    {
        try
        {
            main.CellsU["EventDblClick"].FormulaU = "\"\"";
        }
        catch
        {
        }

        try
        {
            foreach (Visio.Shape child in main.Shapes)
            {
                try
                {
                    child.CellsU["EventDblClick"].FormulaU = "\"\"";
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
    }

    private static Visio.Shape? DetermineLabelShape(Visio.Shape main)
    {
        try
        {
            foreach (Visio.Shape child in main.Shapes)
            {
                if (!string.IsNullOrEmpty(child.Text))
                    return child;
            }
        }
        catch
        {
        }

        return main;
    }

    private static void ConfigureLabelShape(Visio.Shape label, string masterName)
    {
        try
        {
            label.Text = masterName;
        }
        catch
        {
        }

        // Text formatting
        try
        {
            label.CellsU["Char.Size"].FormulaU = $"{_fontSize} pt";
        }
        catch
        {
        }

        try
        {
            label.CellsU["Char.Color"].FormulaU = "RGB(0,0,0)";
        }
        catch
        {
        }

        // Optional font family example:
        if (_fontType != null)
        {
            try
            {
                label.CellsU["Char.Font"].FormulaU = $"FONT(\"{_fontType}\")";
            }
            catch
            {
            }
        }

        // Remove border/fill
        try
        {
            label.CellsU["LinePattern"].FormulaU = "0";
        }
        catch
        {
        }

        try
        {
            label.CellsU["FillPattern"].FormulaU = "0";
        }
        catch
        {
        }

        // Text block layout
        try
        {
            label.CellsU["TxtWidth"].FormulaU = "TEXTWIDTH(TheText)";
            label.CellsU["TxtHeight"].FormulaU = "0.25 in";
            label.CellsU["TxtPinX"].FormulaU = "Width*0.5";
            label.CellsU["TxtPinY"].FormulaU = "-.30";
            label.CellsU["TxtLocPinX"].FormulaU = "TxtWidth*0.5";
            label.CellsU["TxtLocPinY"].FormulaU = "0";
            label.CellsU["Para.HorzAlign"].FormulaU = "1";
            label.CellsU["VerticalAlign"].FormulaU = "0";
        }
        catch
        {
        }

        // Enable text edit on double-click
        try
        {
            label.CellsU["EventDblClick"].FormulaU = "=OPENTEXTWIN()";
        }
        catch
        {
        }
    }

    private static void AddConnectionPoints(Visio.Shape main)
    {
        try
        {
            if (main.SectionExists[(short)Visio.VisSectionIndices.visSectionConnectionPts, 0] == 0)
                main.AddSection((short)Visio.VisSectionIndices.visSectionConnectionPts);

            void Add(string fx, string fy)
            {
                short row = main.AddRow(
                    (short)Visio.VisSectionIndices.visSectionConnectionPts,
                    (short)Visio.VisRowIndices.visRowLast,
                    (short)Visio.VisRowTags.visTagDefault);

                main.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionConnectionPts,
                    row,
                    (short)Visio.VisCellIndices.visCnnctX].FormulaU = fx;
                main.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionConnectionPts,
                    row,
                    (short)Visio.VisCellIndices.visCnnctY].FormulaU = fy;
            }

            Add("-0.05", "0.25*Height");
            Add("-0.05", "0.5*Height");
            Add("-0.05", "0.75*Height");
            Add("0.25*Width", "-0.05");
            Add("0.5*Width", "-0.05");
            Add("0.75*Width", "-0.05");
            Add("0.25*Width", "Height+0.05");
            Add("0.5*Width", "Height+0.05");
            Add("0.75*Width", "Height+0.05");
            Add("Width+0.05", "0.25*Height");
            Add("Width+0.05", "0.5*Height");
            Add("Width+0.05", "0.75*Height");

            main.CellsSRC[
                    (short)Visio.VisSectionIndices.visSectionObject,
                    (short)Visio.VisRowIndices.visRowShapeLayout,
                    (short)Visio.VisCellIndices.visSLOConFixedCode].FormulaU =
                ((int)Visio.VisCellVals.visSLOFixedNoFoldToShape).ToString(CultureInfo.InvariantCulture);
        }
        catch
        {
        }
    }
}

