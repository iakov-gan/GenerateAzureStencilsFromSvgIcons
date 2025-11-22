namespace GenerateAzureStencilsfromSVGIcons.Object
{
    public sealed class AppPathsConfig
    {
        public string IconsRoot { get; }
        public string OutputFolder { get; }

        public AppPathsConfig(string iconsRoot, string outputFolder)
        {
            IconsRoot = iconsRoot ?? throw new ArgumentNullException(nameof(iconsRoot));
            OutputFolder = outputFolder ?? throw new ArgumentNullException(nameof(outputFolder));
        }
    }
}