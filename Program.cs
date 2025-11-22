using GenerateAzureStencilsfromSVGIcons.Object;
using Microsoft.Extensions.Configuration;

namespace GenerateAzureStencilsfromSVGIcons
{
    internal static class Program
    {
        private static IConfigurationRoot BuildConfig()
        {
            return new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();
        }

        private static string ExpandEnv(string? v) =>
            string.IsNullOrWhiteSpace(v) ? string.Empty : Environment.ExpandEnvironmentVariables(v);

        private static AppPathsConfig LoadPaths(IConfiguration cfg)
        {
            // Prefer direct key access
            var iconsRoot = cfg.GetValue<string>("Paths:IconsRoot");
            var outputFolder = cfg.GetValue<string>("Paths:OutputFolder");

            if (string.IsNullOrWhiteSpace(iconsRoot))
                throw new InvalidOperationException("Missing config key Paths:IconsRoot");
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new InvalidOperationException("Missing config key Paths:OutputFolder");

            return new AppPathsConfig(
                ExpandEnv(iconsRoot),
                ExpandEnv(outputFolder));
        }

        private static void Main(string[] args)
        {
            try
            {
                var config = BuildConfig();
                var paths = LoadPaths(config);
                var fontType = config.GetValue<string>("Font:Type");
                var fontSize = Convert.ToInt16(config.GetValue<int?>("Font:Size") ?? 10);

                Console.WriteLine("Azure SVG -> Visio stencil generator");
                Console.WriteLine($"Icons root   : {paths.IconsRoot}");
                Console.WriteLine($"Output folder: {paths.OutputFolder}");
                Console.WriteLine();

                Directory.CreateDirectory(paths.OutputFolder);

                var discovery = new SvgDiscoveryService();
                var groups = discovery.DiscoverGroups(paths.IconsRoot);

                var formatter = new NameFormatter();
                var generator = new VisioStencilGenerator(paths, formatter, fontType, fontSize);
                generator.GenerateStencils(groups);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR: " + ex.Message);
            }
        }
    }
}
