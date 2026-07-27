using Syncfusion.Presentation;

namespace Convert_PPTX_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open("Data/Input.pptx"))
            {
                //Save the PowerPoint Presentation as a Markdown file.
                presentation.Save("Output/PPTXtoMd.md");
            }
        }
    }
}