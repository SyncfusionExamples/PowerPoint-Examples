using Syncfusion.Presentation;
using System.IO;
using System.Web.Mvc;

namespace Read_and_edit_PowerPoint_presentation.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult CreatePowerPoint()
        {
            using (FileStream fileStreamPath = new FileStream(Server.MapPath("~/App_Start/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
                {
                    //Gets the first slide from the PowerPoint presentation
                    ISlide slide = pptxDoc.Slides[0];
                    //Gets the first shape of the slide
                    IShape shape = slide.Shapes[0] as IShape;
                    //Change the text of the shape
                    if (shape.TextBody.Text == "Company History")
                        shape.TextBody.Text = "Company Profile";
                    //Save the PowerPoint Presentation as stream
                    MemoryStream pptxStream = new MemoryStream();
                    pptxDoc.Save(pptxStream);
                    pptxStream.Position = 0;
                    //Download Powerpoint document in the browser.
                    return File(pptxStream, "application/powerpoint", "Result.pptx");
                }
            }
        }
    }
}