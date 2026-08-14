using Syncfusion.Presentation;
using System.Text;
using System.IO;

namespace Encoding_As_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Input.pptx")))
            {
                // Set the encoding for the Markdown file.
                presentation.MdSaveOptions.Encoding = Encoding.ASCII;
                // Save the PowerPoint Presentation as a Markdown file.
                presentation.Save(Path.GetFullPath("Output.md"));
            }
        }
    }
}
