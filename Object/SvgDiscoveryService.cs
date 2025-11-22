namespace GenerateAzureStencilsfromSVGIcons.Object
{
    public sealed class SvgDiscoveryService
    {
        public IEnumerable<IGrouping<string?, string>> DiscoverGroups(string iconsRoot)
        {
            if (!Directory.Exists(iconsRoot))
                throw new DirectoryNotFoundException($"Icons root does not exist: {iconsRoot}");

            var allSvgFiles = Directory
                .EnumerateFiles(iconsRoot, "*.svg", SearchOption.AllDirectories)
                .ToList();

            if (!allSvgFiles.Any())
                throw new InvalidOperationException($"No SVG files found under {iconsRoot}");

            return allSvgFiles
                .GroupBy(Path.GetDirectoryName)
                .ToList();
        }
    }
}