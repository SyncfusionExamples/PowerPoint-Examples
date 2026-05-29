using Syncfusion.Presentation;

namespace Convert_Presentation_to_Markdown
{
    class Program
    {

        static void Main(string[] args)
        {
            //Open the file as a Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Load the file stream
                using (IPresentation presentation = Presentation.Open(fileStream))
                {
                    //Save as a PPTX document into the Markdown FileStream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/MarkdownToPPTX.md"), FileMode.Create))
                    {
                        presentation.Save(outputStream);
                    }
                }
            }
        }
    }
}
