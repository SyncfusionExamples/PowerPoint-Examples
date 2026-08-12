using Syncfusion.Presentation;
using System.IO;
using SkiaSharp;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_Presentation_to_Image.Data
{
    public class PowerPointService
    {
        public MemoryStream ConvertPPTXtoImage()
        {
            //Opens an existing PowerPoint presentation from the wwwroot folder.
            using (IPresentation pptxDoc = Presentation.Open(Path.Combine("wwwroot", "Input.pptx")))
            {
                //Initialize the PresentationRenderer to perform image conversion.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert PowerPoint slide to image as stream.
                using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                {
                    //Save the converted image file to MemoryStream.
                    MemoryStream outputStream = new MemoryStream();
                    stream.CopyTo(outputStream);
                    outputStream.Position = 0;
                    //Return the image stream for download in the browser.
                    return outputStream;
                }
            }
        }
    }
}