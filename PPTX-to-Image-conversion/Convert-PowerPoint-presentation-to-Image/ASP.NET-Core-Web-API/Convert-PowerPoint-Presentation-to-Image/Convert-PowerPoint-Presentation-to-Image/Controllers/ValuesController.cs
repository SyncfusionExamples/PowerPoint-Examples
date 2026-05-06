using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
using static System.Runtime.InteropServices.JavaScript.JSType;


namespace Convert_PowerPoint_Presentation_to_Image.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/PPTXToImage")]
        public IActionResult ConvertPPTXToImage()
        {
            try
            {
                var fileDownloadName = "Output.jpeg";
                const string contentType = "image/jpeg";
                var stream = ConvertPresentationToImage();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                // Log or handle the exception
                return BadRequest("Error occurred while converting PowerPoint Presentation to Image: " + ex.Message);
            }
        }
        public static Stream ConvertPresentationToImage()
        {
            IPresentation pptxDoc = Presentation.Open("Data/Input.pptx");
            //Initialize the PresentationRenderer to perform image conversion.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Convert the first page of the Word document into an image.
            Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
            //close the word document.
            pptxDoc.Close();
            //Reset the stream position.
            stream.Position = 0;
            //Save the image file.
            return stream;
        }
    }
}
