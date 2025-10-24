using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;


namespace Convert_PowerPoint_Presentation_to_PDF.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/ConvertToPdf")]
        public IActionResult ConvertToPdf()
        {
            try
            {
                var fileDownloadName = "Converted.pdf";
                const string contentType = "application/pdf";
                var stream = ConvertPresentationToPdf();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                return BadRequest("Error occurred while converting PowerPoint to PDF: " + ex.Message);
            }
        }
        public static MemoryStream ConvertPresentationToPdf()
        {
            //Open the existing PowerPoint presentation with loaded stream.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
            {

                //Convert the PowerPoint presentation to PDF document.
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
