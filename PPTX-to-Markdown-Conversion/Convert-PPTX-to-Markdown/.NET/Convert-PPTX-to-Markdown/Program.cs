using Syncfusion.Presentation;
using System.IO;

namespace Convert_PPTX_To_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
            {
                //Save the PowerPoint Presentation as a Markdown file.
                presentation.Save(Path.GetFullPath("Output/PPTXtoMd.md"));
            }
        }
    }
}