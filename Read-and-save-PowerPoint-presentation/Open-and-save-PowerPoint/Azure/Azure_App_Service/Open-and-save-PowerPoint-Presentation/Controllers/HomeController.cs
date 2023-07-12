using Microsoft.AspNetCore.Mvc;
using Open_and_save_PowerPoint_Presentation.Models;
using Syncfusion.Presentation;
using System.Diagnostics;

namespace Open_and_save_PowerPoint_Presentation.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public ActionResult OpenAndSavePowerPoint()
        {
            string pptxPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data/Template.pptx");

            using (FileStream fileStreamPath = new FileStream(pptxPath, FileMode.Open, FileAccess.Read))
            {
                //Open an existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
                {
                    //Get the first slide from the PowerPoint presentation.
                    ISlide slide = pptxDoc.Slides[0];
                    //Get the first shape of the slide.
                    IShape shape = slide.Shapes[0] as IShape;
                    //Change the text of the shape.
                    if (shape.TextBody.Text == "Company History")
                        shape.TextBody.Text = "Company Profile";
                    //Save the PowerPoint Presentation as stream.
                    MemoryStream pptxStream = new();
                    pptxDoc.Save(pptxStream);
                    pptxStream.Position = 0;
                    //Download Powerpoint document in the browser.
                    return File(pptxStream, "application/powerpoint", "Result.pptx");
                }
            }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}