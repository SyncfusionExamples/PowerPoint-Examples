
using Syncfusion.Presentation;
using System.IO;

namespace Use_Thematic_Break_As_ContentBreak
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create load options.
            LoadOptions loadOptions = new LoadOptions();
            // Specify the format type as Markdown.
            loadOptions.FormatType = FormatType.Markdown;
            // Initialize Markdown import settings for the LoadOptions instance.
            loadOptions.MdImportSettings = new Syncfusion.Office.Markdown.MdImportSettings();
            // Set UseThematicBreakAsContentBreak to split slides based on thematic breaks.
            loadOptions.MdImportSettings.UseThematicBreakAsContentBreak = true;
            // Open the Markdown file with load options.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Data/Input.md"), loadOptions))
            {
                // Save as a PowerPoint document.
                presentation.Save(Path.GetFullPath("Output/MarkdownToPPTX.pptx"));
            }
        }
    }
}
