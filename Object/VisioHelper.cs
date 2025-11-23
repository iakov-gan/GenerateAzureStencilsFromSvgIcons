namespace GenerateAzureStencilsFromSvgIcons.Object;

using System.Runtime.InteropServices;
using Visio = Microsoft.Office.Interop.Visio;

public static class VisioHelper
{
    public static void SafeQuit(Visio.Application? visioApp)
    {
        if (visioApp != null)
        {
            try
            {
                // Close all documents first (stencils are documents too)
                while (visioApp.Documents.Count > 0)
                {
                    var doc = visioApp.Documents[1];  // Index starts at 1 in Visio COM
                    doc.Saved = true;  // Mark as saved to avoid save prompts (adjust if you need actual saving)
                    doc.Close();
                    Marshal.FinalReleaseComObject(doc);  // Release each doc reference
                }

                // Now quit the app
                visioApp.Quit();

                // Release the main app reference
                Marshal.FinalReleaseComObject(visioApp);
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Exception in VisioHelper.SafeQuit; Visio may not have closed cleanly: {exception.Message}");
            }
            finally
            {
                visioApp = null;

                // Force GC to clean up any lingering COM refs
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
