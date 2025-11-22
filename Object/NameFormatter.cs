using System.Globalization;
using System.Text.RegularExpressions;

namespace GenerateAzureStencilsfromSVGIcons.Object
{
    public sealed class NameFormatter
    {
        private static readonly HashSet<string> Acronyms =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "AI", "API", "IOT", "SQL", "VM", "VPN", "VNET", "DNS", "HTTP", "HTTPS"
            };

        private static readonly TextInfo InvariantTextInfo =
            CultureInfo.InvariantCulture.TextInfo;

        public string ToFolderTitleCase(string folderName)
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

        public string GetMasterNameFromFile(string fileName)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName) ?? fileName;

            baseName = Regex.Replace(
                baseName,
                @"^\d+-icon-service-",
                string.Empty,
                RegexOptions.IgnoreCase);

            baseName = baseName.Replace("-", " ");

            var tokens = Regex.Split(baseName, @"(\s+)");

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
                            trimmed[1..].ToLowerInvariant();
                    }
                    else
                    {
                        tokens[i] = trimmed.ToUpperInvariant();
                    }
                }
            }

            return string.Concat(tokens).Trim();
        }
    }
}
