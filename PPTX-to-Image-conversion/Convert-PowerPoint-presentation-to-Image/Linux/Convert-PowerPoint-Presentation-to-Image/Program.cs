using System;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
using static System.Collections.Specialized.BitVector32;


namespace Convert_PowerPoint_Presentation_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStreamInput = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                    {
                        //Reset the stream position.
                        stream.Position = 0;
                        //Create FileStream to save the image file.
                        using (FileStream outputStream = new FileStream("PPTXtoImage.Jpeg", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            //Save the image file.
                            stream.CopyTo(outputStream);
                        }
                    }
                }
            }
        }
    }
}
