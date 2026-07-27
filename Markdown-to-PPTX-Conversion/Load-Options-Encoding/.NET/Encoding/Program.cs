
using Syncfusion.Presentation;
using System.Text;

namespace Encoding_as_PPTX
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
            //Set the encoding for the Markdown file.
            loadOptions.MdImportSettings.Encoding = Encoding.UTF8;
            // Open the Markdown file with load options.
            using (IPresentation presentation = Presentation.Open("Data/Input.md", loadOptions))
            {
                //Save as a PowerPoint document.
                presentation.Save("Output/MarkdownToPPTX.pptx");
            }
        }
    }
}
