using Convert_PowerPoint_Presentation_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
using SkiaSharp;

namespace Convert_PowerPoint_Presentation_to_Image.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// Convert Presentation to image
        /// </summary>
        /// <param name="button"></param>
        /// <returns></returns>
        public ActionResult ConvertPPTXtoImage(string button)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath("Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    stream.Position = 0;
                    //Download image in the browser.
                    return File(stream, "application/jpeg", "PPTXtoImage.Jpeg");
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