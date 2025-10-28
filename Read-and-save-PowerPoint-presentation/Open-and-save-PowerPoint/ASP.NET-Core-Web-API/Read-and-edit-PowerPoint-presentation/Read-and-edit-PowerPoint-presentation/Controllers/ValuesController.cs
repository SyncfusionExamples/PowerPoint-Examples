using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Presentation;

namespace Read_and_edit_PowerPoint_presentation.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [HttpGet]
        [Route("api/PowerPoint")]
        public IActionResult DownloadPresentation()
        {
            try
            {
                var fileDownloadName = "Output.pptx";
                const string contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                var stream = OpenandSavePresentation();
                stream.Position = 0;
                return File(stream, contentType, fileDownloadName);
            }
            catch (Exception ex)
            {
                // Log or handle the exception
                return BadRequest("Error occurred while creating PowerPoint Presentation: " + ex.Message);
            }
        }
        public static MemoryStream OpenandSavePresentation()
        {
            IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
            //Get the first slide from the PowerPoint presentation.
            ISlide slide = pptxDoc.Slides[0];
            //Get the first shape of the slide.
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape.
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            // Save the PowerPoint Presentation as stream
            MemoryStream stream = new MemoryStream();
            pptxDoc.Save(stream);
            pptxDoc.Close();
            stream.Position = 0;
            return stream;
        }
    }
}
