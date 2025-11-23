namespace GenerateAzureStencilsFromSvgIcons.Service;

using System.Text.RegularExpressions;
using System.IO.Compression;

public class AzureIconsDownloader
{
    private readonly string _pageUrl;

    public AzureIconsDownloader(string pageUrl = "https://learn.microsoft.com/en-us/azure/architecture/icons/")
    {
        _pageUrl = pageUrl;
    }

    public async Task DownloadAsync(string destinationPath)
    {
        Console.WriteLine("Fetching the Microsoft Azure Architecture Icons download page...");
        Console.WriteLine($"Page: {_pageUrl}");

        string downloadUrl = null;

        try
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

                string html = null;
                try
                {
                    html = await client.GetStringAsync(_pageUrl);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.InnerException.Message);
                }

                // Regex to find the <a> tag with "Download SVG icons" and capture the href
                string pattern = @"<a\s+[^>]*href=""(https?://[^""]+\.zip)""[^>]*>Download\s+SVG\s+icons</a>";
                Match match = Regex.Match(html, pattern, RegexOptions.IgnoreCase);

                if (match.Success && match.Groups.Count > 1)
                {
                    downloadUrl = match.Groups[1].Value;
                    Console.WriteLine($"Found download URL: {downloadUrl}");
                }
                else
                {
                    Console.WriteLine("Could not find the download link in the page HTML.");
                    Console.WriteLine("Falling back to a known direct link (may not be the latest).");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error fetching page: {ex.Message}");
            Console.WriteLine("Using fallback direct link.");
        }

        if (string.IsNullOrEmpty(downloadUrl))
        {
            Console.WriteLine("No download URL found. Exiting.");
            Console.ReadKey();
            return;
        }

        string fileName = Path.GetFileName(new Uri(downloadUrl).AbsolutePath);
        if (string.IsNullOrEmpty(fileName))
        {
            Console.WriteLine("Could not derive a file name from the download URL. Exiting.");
            Console.ReadKey();
            return;
        }

        // REPLACED: destination path handling now builds a date subfolder and cleans if exists.
        string baseFolder;
        if (string.IsNullOrWhiteSpace(destinationPath))
        {
            var downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads");
            baseFolder = downloadsFolder;
        }
        else
        {
            // Normalize: if a file path with extension supplied, use its directory; else treat as directory.
            if (!string.IsNullOrWhiteSpace(Path.GetExtension(destinationPath)) && File.Exists(destinationPath) ||
                !string.IsNullOrWhiteSpace(Path.GetExtension(destinationPath)) &&
                !destinationPath.EndsWith(Path.DirectorySeparatorChar) &&
                !destinationPath.EndsWith(Path.AltDirectorySeparatorChar))
            {
                var dir = Path.GetDirectoryName(destinationPath)!;
                baseFolder = string.IsNullOrEmpty(dir) ? Directory.GetCurrentDirectory() : dir;
            }
            else
            {
                baseFolder = destinationPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
        }

        if (!Directory.Exists(baseFolder))
            Directory.CreateDirectory(baseFolder);

        string dateFolder = Path.Combine(baseFolder, DateTime.Now.ToString("yyyy-MM-dd"));
        if (Directory.Exists(dateFolder))
        {
            Console.WriteLine($"Cleaning existing dated folder: {dateFolder}");
            Directory.Delete(dateFolder, true);
        }

        Directory.CreateDirectory(dateFolder);

        destinationPath = Path.Combine(dateFolder, fileName);
        Console.WriteLine($"Downloading to dated folder: {destinationPath}");

        try
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

                using (HttpResponseMessage response = await client.GetAsync(downloadUrl))
                {
                    response.EnsureSuccessStatusCode();

                    using (Stream contentStream = await response.Content.ReadAsStreamAsync(),
                           fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write,
                               FileShare.None))
                    {
                        await contentStream.CopyToAsync(fileStream);
                    }
                }
            }

            Console.WriteLine("Download completed successfully!");
            Console.WriteLine($"File saved as: {destinationPath}");

            // NEW: extract zip into same dated folder.
            Console.WriteLine("Extracting ZIP contents...");
            using (ZipArchive archive = ZipFile.OpenRead(destinationPath))
            {
                foreach (var entry in archive.Entries)
                {
                    if (string.IsNullOrEmpty(entry.Name))
                        continue; // directory entry
                    string targetPath = Path.Combine(dateFolder, entry.FullName);
                    string? entryDir = Path.GetDirectoryName(targetPath);
                    if (!string.IsNullOrEmpty(entryDir) && !Directory.Exists(entryDir))
                        Directory.CreateDirectory(entryDir);
                    entry.ExtractToFile(targetPath, true);
                }
            }

            Console.WriteLine("Extraction completed.");

            Console.WriteLine("\nNote: This ZIP contains SVG icons which can be imported into Visio.");
            Console.WriteLine("For native Visio stencils (.vssx), consider community repositories like:");
            Console.WriteLine("- https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio");
            Console.WriteLine("- https://github.com/David-Summers/Azure-Design");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error downloading or extracting: {ex.Message}");
            Console.WriteLine(
                "Verify the URL on the page: https://learn.microsoft.com/en-us/azure/architecture/icons/");
        }

        // Console.WriteLine("\nPress any key to exit...");
        // Console.ReadKey();
    }
}