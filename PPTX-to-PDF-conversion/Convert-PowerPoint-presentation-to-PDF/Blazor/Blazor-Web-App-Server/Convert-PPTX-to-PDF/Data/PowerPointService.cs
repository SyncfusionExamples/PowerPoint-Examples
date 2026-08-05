using Syncfusion.Presentation;
using Syncfusion.Pdf;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_Presentation_to_PDF.Data
{
    public class PowerPointService
    {
        public MemoryStream ConvertPPTXtoPDF()
        {
            // Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open("wwwroot/Input.pptx"))
            {
                // Convert the PowerPoint presentation to PDF document.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    // Save the converted PDF document to a MemoryStream.
                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        pdfDocument.Save(pdfStream);
                        // Reset stream position before returning it to the browser.
                        pdfStream.Position = 0;
                        // Return the PDF document for download in the browser.
                        return pdfStream;
                    }
                }
            }
        }
    }
}