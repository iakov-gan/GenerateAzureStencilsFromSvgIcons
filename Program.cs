using System.Globalization;
using Visio = Microsoft.Office.Interop.Visio;

namespace AzureSvgToVisioStencils
{
    internal static class Program
    {
        // CONFIGURE THESE AS NEEDED
        private static readonly string IconsRoot =
            @"H:\Free Software\Visio\Icons\Azure\Azure_Public_Service_Icons_V23\Azure_Public_Service_Icons\Icons";

        private static readonly string OutputFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                @"My Shapes\Azure Icons");

        // Acronyms to keep upper-case in master names
        private static readonly HashSet<string> Acronyms =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "AI", "API", "IOT", "SQL", "VM", "VPN", "VNET", "DNS", "HTTP", "HTTPS"
            };

        private static readonly TextInfo InvariantTextInfo =
            CultureInfo.InvariantCulture.TextInfo;

        private static void Main(string[] args)
        {
            Console.WriteLine("Azure SVG -> Visio stencil generator (C#)");
            Console.WriteLine($"Icons root   : {IconsRoot}");
            Console.WriteLine($"Output folder: {OutputFolder}");
            Console.WriteLine();

            if (!Directory.Exists(IconsRoot))
            {
                Console.Error.WriteLine($"ERROR: Icons root does not exist: {IconsRoot}");
                return;
            }

            Directory.CreateDirectory(OutputFolder);

            // Gather all SVGs by folder
            var allSvgFiles = Directory
                .EnumerateFiles(IconsRoot, "*.svg", SearchOption.AllDirectories)
                .ToList();

            if (!allSvgFiles.Any())
            {
                Console.Error.WriteLine($"ERROR: No SVG files found under {IconsRoot}");
                return;
            }

            Console.WriteLine($"Found {allSvgFiles.Count} SVG files.");
            Console.WriteLine();

            var groups = allSvgFiles
                .GroupBy(Path.GetDirectoryName)
                .ToList();

            Console.WriteLine($"Category folders (stencils) to be created: {groups.Count}");
            Console.WriteLine();

            Visio.Application visioApp = null;

            try
            {
                visioApp = new Visio.Application
                {
                    Visible = false
                };

                int groupIndex = 0;

                foreach (var group in groups)
                {
                    groupIndex++;

                    string folderPath = group.Key ?? IconsRoot;
                    string rawCategory = Path.GetFileName(folderPath);
                    if (string.IsNullOrWhiteSpace(rawCategory))
                        rawCategory = "(Root)";

                    string prettyCategory = ToFolderTitleCase(rawCategory);
                    string stencilName = $"Azure-{prettyCategory}.vssx";
                    string stencilPath = Path.Combine(OutputFolder, stencilName);

                    Console.WriteLine($"[{groupIndex} / {groups.Count}] Folder     : '{rawCategory}'");
                    Console.WriteLine($"[{groupIndex} / {groups.Count}] Stencil    : '{stencilName}'");

                    var svgFiles = group.ToList();
                    if (!svgFiles.Any())
                    {
                        Console.WriteLine("  (No SVGs in this folder, skipping.)");
                        Console.WriteLine();
                        continue;
                    }

                    // Create a new stencil document
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

                        string masterName = GetMasterNameFromFile(fileName);

                        // Create a new master
                        Visio.Master master = masters.Add();
                        master.Name = masterName;

                        // Import the SVG into the master
                        Visio.Shape importedShape = master.Import(svgPath);

                        // Normalize icon size
                        double oldWidthIU = importedShape.CellsU["Width"].ResultIU;
                        importedShape.CellsU["Width"].FormulaU = "0.5 in";
                        // keep aspect ratio
                        importedShape.CellsU["Height"].FormulaU =
                            (importedShape.CellsU["Height"].ResultIU * (0.5 / oldWidthIU)).ToString(CultureInfo
                                .InvariantCulture) + " in";

                        // Try to find a top-level group; if none, use importedShape
                        Visio.Shape targetShape = FindTopLevelGroup(master) ?? importedShape;

                        // Configure text, double-click, selection, connection points
                        InitializeMasterShape(targetShape, masterName);
                    }

                    stencilDoc.SaveAs(stencilPath);
                    stencilDoc.Close();

                    Console.WriteLine($"  Completed stencil: {stencilPath}");
                    Console.WriteLine();
                }

                Console.WriteLine("All stencils created.");
                Console.WriteLine("Open Visio and check: More Shapes -> My Shapes -> Azure-*.vssx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR: " + ex.Message);
                Console.Error.WriteLine(ex);
            }
            finally
            {
                if (visioApp != null)
                {
                    try
                    {
                        visioApp.Quit();
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }
        }

        /// <summary>
        /// Title-cases the folder name for stencil naming.
        /// </summary>
        private static string ToFolderTitleCase(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
                return "(Root)";

            var parts = folderName
                .Split(' ')
                .Select(p =>
                {
                    if (string.IsNullOrWhiteSpace(p))
                        return p;
                    return InvariantTextInfo.ToTitleCase(p.ToLowerInvariant());
                });

            return string.Join(" ", parts);
        }

        /// <summary>
        /// Clean master names: strip ###-icon-service-, replace dashes with spaces,
        /// title-case with acronym exceptions.
        /// </summary>
        private static string GetMasterNameFromFile(string fileName)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName) ?? fileName;

            // Strip leading "###-icon-service-"
            baseName = System.Text.RegularExpressions.Regex.Replace(
                baseName,
                @"^\d+-icon-service-",
                string.Empty,
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Replace dashes with spaces
            baseName = baseName.Replace("-", " ");

            // Split on whitespace, preserving spaces between tokens
            var tokens = System.Text.RegularExpressions.Regex.Split(baseName, @"(\s+)");

            for (int i = 0; i < tokens.Length; i++)
            {
                string token = tokens[i];
                string trimmed = token.Trim();

                if (trimmed.Length == 0)
                    continue;

                if (Acronyms.Contains(trimmed))
                {
                    tokens[i] = trimmed.ToUpperInvariant();
                }
                else
                {
                    if (trimmed.Length > 1)
                    {
                        tokens[i] =
                            char.ToUpperInvariant(trimmed[0]) +
                            trimmed.Substring(1).ToLowerInvariant();
                    }
                    else
                    {
                        tokens[i] = trimmed.ToUpperInvariant();
                    }
                }
            }

            string result = string.Concat(tokens).Trim();
            return result;
        }

        /// <summary>
        /// Try to find the top-level group shape in a master, if any.
        /// </summary>
        private static Visio.Shape FindTopLevelGroup(Visio.Master master)
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

    // Ensure group-only selection (if group)
    try
    {
        if (main.Type == (short)Visio.VisShapeTypes.visTypeGroup)
        {
            var selectModeCell = main.CellsSRC[
                (short)Visio.VisSectionIndices.visSectionObject,
                (short)Visio.VisRowIndices.visRowGroup,
                (short)Visio.VisCellIndices.visGroupSelectMode];
            selectModeCell.FormulaU = "0"; // group-only
        }
    }
    catch { }

    // Clear dbl-click on main
    try { main.CellsU["EventDblClick"].FormulaU = "\"\""; } catch { }

    // Clear dbl-click on children
    try
    {
        foreach (Visio.Shape child in main.Shapes)
        {
            try { child.CellsU["EventDblClick"].FormulaU = "\"\""; } catch { }
        }
    }
    catch { }

    // Do not use group text
    try { main.Text = string.Empty; } catch { }

    Visio.Shape label = null;

    // Try to reuse existing text-bearing child if any
    try
    {
        foreach (Visio.Shape child in main.Shapes)
        {
            if (!string.IsNullOrEmpty(child.Text))
            {
                label = child;
                break;
            }
        }

        // If no label child exists, just use main as the label host
        if (label == null)
        {
            label = main;
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine(
            $"ERROR choosing label shape for master '{masterName}': {ex.GetType().FullName}: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
    }

    if (label != null)
    {
        // Default text
        try { label.Text = masterName; } catch { }

        // Text formatting
        try { label.CellsU["Char.Size"].FormulaU = "9 pt"; } catch { }
        try { label.CellsU["Char.Color"].FormulaU = "RGB(0,0,0)"; } catch { }

        // Hide border/fill where possible
        try { label.CellsU["LinePattern"].FormulaU = "0"; } catch { }
        try { label.CellsU["FillPattern"].FormulaU = "0"; } catch { }

        // Text block layout: bottom-center anchor, grow downward
        try
        {
            // Make text area wider than icon
            label.CellsU["TxtWidth"].FormulaU = "TEXTWIDTH(TheText)";

            // Initial text height
            label.CellsU["TxtHeight"].FormulaU = "0.25 in";

            // Pin at bottom-center of the shape
            label.CellsU["TxtPinX"].FormulaU = "Width*0.5";
            label.CellsU["TxtPinY"].FormulaU = "-.30"; // bottom of shape

            // Local pin at bottom-center so growth is downward
            label.CellsU["TxtLocPinX"].FormulaU = "TxtWidth*0.5";
            label.CellsU["TxtLocPinY"].FormulaU = "0";

            // Center text horizontally
            label.CellsU["Para.HorzAlign"].FormulaU = "1";
            
            // Vertically align text to top of text block
            // label.CellsU["Para.VAlign"].FormulaU = "0";
            label.CellsU["VerticalAlign"].FormulaU = "0";  
        }
        catch { }

        // Double-click on label opens text edit
        try { label.CellsU["EventDblClick"].FormulaU = "=OPENTEXTWIN()"; } catch { }
    }

    // Connection points on the main shape
    try
    {
        if (main.SectionExists[(short)Visio.VisSectionIndices.visSectionConnectionPts, 0] == 0)
            main.AddSection((short)Visio.VisSectionIndices.visSectionConnectionPts);

        void AddConnPt(string fx, string fy)
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

        AddConnPt("0", "0.25*Height");
        AddConnPt("0", "0.5*Height");
        AddConnPt("0", "0.75*Height");
        AddConnPt("0.25*Width", "0");
        AddConnPt("0.5*Width", "0");
        AddConnPt("0.75*Width", "0");
        AddConnPt("0.25*Width", "Height");
        AddConnPt("0.5*Width", "Height");
        AddConnPt("0.75*Width", "Height");
        AddConnPt("Width", "0.25*Height");
        AddConnPt("Width", "0.5*Height");
        AddConnPt("Width", "0.75*Height");

        main.CellsSRC[
            (short)Visio.VisSectionIndices.visSectionObject,
            (short)Visio.VisRowIndices.visRowShapeLayout,
            (short)Visio.VisCellIndices.visSLOConFixedCode].FormulaU =
            ((int)Visio.VisCellVals.visSLOFixedNoFoldToShape)
            .ToString(CultureInfo.InvariantCulture);
    }
    catch { }
}
    }
}
