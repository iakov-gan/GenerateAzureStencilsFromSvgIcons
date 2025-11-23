namespace GenerateAzureStencilsFromSvgIcons.Model
{
    public sealed class AppPathsConfigModel
    {
        public string IconsRoot { get; }
        public string OutputFolder { get; }

        public AppPathsConfigModel(string iconsRoot, string outputFolder)
        {
            IconsRoot = iconsRoot ?? throw new ArgumentNullException(nameof(iconsRoot));
            OutputFolder = outputFolder ?? throw new ArgumentNullException(nameof(outputFolder));
        }
    }
}