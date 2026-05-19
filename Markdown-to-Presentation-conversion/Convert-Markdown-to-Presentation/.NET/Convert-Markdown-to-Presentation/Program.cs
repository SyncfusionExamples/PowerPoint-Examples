using Syncfusion.Presentation;

namespace Convert_Markdown_to_Presentation
{
    class Program
    {

        static void Main(string[] args)
        {
            //Open the file as a Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.md"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream.
                using (IPresentation presentation = Presentation.Open(fileStream))
                {
                    //Save as a Markdown document into the PPTX FileStream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/MarkdownToPPTX.pptx"), FileMode.Create))
                    {
                        presentation.Save(outputStream);
                    }
                }
            }
        }
    }
}
