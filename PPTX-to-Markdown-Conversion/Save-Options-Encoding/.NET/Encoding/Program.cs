using Syncfusion.Presentation;
using System.Text;

namespace Encoding_as_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open("Input.pptx"))
            {
                // Set the encoding for the Markdown file.
                presentation.MdSaveOptions.Encoding = Encoding.ASCII;
                // Save the PowerPoint Presentation as a Markdown file.
                presentation.Save("Output.md");
            }
        }
    }
}
