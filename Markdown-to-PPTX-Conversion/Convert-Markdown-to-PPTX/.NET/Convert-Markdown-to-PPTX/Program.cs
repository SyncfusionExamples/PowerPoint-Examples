using Syncfusion.Presentation;
using System.IO;

namespace Convert_Markdown_To_PowerPoint
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Markdown file.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Input.md")))
            {
                //Save as a PowerPoint document.
                presentation.Save(Path.GetFullPath(@"Output/MarkdownToPPTX.pptx"));
            }
        }
    }
}


