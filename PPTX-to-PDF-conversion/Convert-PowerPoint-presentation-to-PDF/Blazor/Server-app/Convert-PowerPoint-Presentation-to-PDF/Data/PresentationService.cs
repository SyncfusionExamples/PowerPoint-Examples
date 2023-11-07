using Syncfusion.Presentation;
using Syncfusion.Pdf;
using System.IO;
using SkiaSharp;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_Presentation_to_PDF.Data
{
    public class PresentationService
    {
        public MemoryStream ConvertPPTXtoPDF()
        {
            //Open the file as Stream
            using (FileStream sourceStreamPath = new FileStream(@"wwwroot/Input.pptx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(sourceStreamPath))
                {
                    //Convert the PowerPoint document to PDF document.
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        //Create the MemoryStream to save the converted PDF.      
                        MemoryStream pdfStream = new MemoryStream();
                        //Save the converted PDF document to MemoryStream.
                        pdfDocument.Save(pdfStream);
                        pdfStream.Position = 0;

                        //Download PDF document in the browser.
                        return pdfStream;
                    }
                }
            }                                           
        }
    }
}
