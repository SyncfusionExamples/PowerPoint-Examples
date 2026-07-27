using Syncfusion.Presentation;
using System.IO;

namespace Convert_Markdown_to_PowerPoint
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Markdown file.
            using (IPresentation presentation = Presentation.Open("Data/Input.md"))
            {
                //Save as a PowerPoint document.
                presentation.Save("Output/MarkdownToPPTX.pptx");
            }
        }
    }
}


