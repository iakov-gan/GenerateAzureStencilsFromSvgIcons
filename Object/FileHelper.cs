namespace GenerateAzureStencilsFromSvgIcons.Object;

using System.IO;

public static class FileHelper
{
    public static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path))
        {
            try
            {
                Directory.Delete(path, recursive: true);
            }
            catch (DirectoryNotFoundException)
            {
                // Folder disappeared between Exists check and Delete; ignore.
            }
        }
    }
}
