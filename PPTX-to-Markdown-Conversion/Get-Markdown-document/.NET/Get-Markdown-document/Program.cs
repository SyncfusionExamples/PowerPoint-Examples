using Syncfusion.Office.Markdown;
using Syncfusion.Presentation;
using System.IO;
using System.Text;

namespace Get_Markdown_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
            {
                //Convert the PowerPoint Presentation to Markdown.
                MarkdownDocument markdownDocument = presentation.GetMarkdownDocument();
                //Save or process the Markdown document as needed.
                markdownDocument.Save(Path.GetFullPath("Output/Output.md"));
                //Dispose the Markdown document.
                markdownDocument.Dispose();
            }
        }
    }
}
