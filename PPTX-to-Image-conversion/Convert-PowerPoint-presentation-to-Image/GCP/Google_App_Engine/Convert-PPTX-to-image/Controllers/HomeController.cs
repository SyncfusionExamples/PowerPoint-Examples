using Convert_PPTX_to_image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Drawing;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Convert_PPTX_to_image.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        public ActionResult ConvertPPTXToImage()
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                //Open the existing PPTX document.
                using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    stream = (MemoryStream)pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                    stream.Position = 0;
                    //Download image file in the browser.
                    return File(stream, "image/jpeg", "PPTXToimage_Page1.jpeg");
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
            }
            return View();
        }

        public IActionResult Index()
        {
            return View();
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