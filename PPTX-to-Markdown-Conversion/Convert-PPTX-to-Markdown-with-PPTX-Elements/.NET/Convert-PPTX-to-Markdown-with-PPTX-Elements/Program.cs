using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;

namespace Convert_PPTX_To_Markdown_with_PPTX_Elements
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
            {
                //Initialize the PresentationRenderer to preserve PowerPoint elements as images.
                presentation.PresentationRenderer = new PresentationRenderer();
                //Save the PowerPoint Presentation as a Markdown file.
                presentation.Save(Path.GetFullPath("Output/Output.md"));
            }
        }
    }
}