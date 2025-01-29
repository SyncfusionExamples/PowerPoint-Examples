using Syncfusion.Presentation;
using System.IO;
using SkiaSharp;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_Presentation_to_Image.Data
{
    public class PowerPointService
    {
        public MemoryStream ConvertPPTXToImage()
        {
            //Open the file as Stream.
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/Input.pptx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(sourceStreamPath))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                    {
                        //Save the converted image file to MemoryStream.
                        MemoryStream Stream = new MemoryStream();
                        stream.CopyTo(Stream);
                        Stream.Position = 0;
                        //Download image file in the browser.
                        return Stream;
                    }
                }
            }                                           
        }
    }
}