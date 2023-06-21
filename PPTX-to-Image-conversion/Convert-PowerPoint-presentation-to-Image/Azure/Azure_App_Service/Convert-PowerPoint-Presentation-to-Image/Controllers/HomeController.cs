using Convert_PowerPoint_Presentation_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.Diagnostics;
using System.IO;

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
            if (button == null)
                return View("Index");

            if (Request.Form.Files != null)
            {
                if (Request.Form.Files.Count == 0)
                {
                    ViewBag.Message = string.Format("Browse a PowerPoint Presentation and then click the button to convert as a image");
                    return View("Index");
                }
                // Gets the extension from file.
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();
                // Compares extension with supported extensions.
                if (extension == ".pptx")
                {
                    MemoryStream memoryStream = new MemoryStream();
                    Request.Form.Files[0].CopyTo(memoryStream);
                    try
                    {
                        //Open the existing PowerPoint presentation with loaded stream.
                        using (IPresentation pptxDoc = Presentation.Open(memoryStream))
                        {

                            //Initialize the PresentationRenderer to perform image conversion.
                            pptxDoc.PresentationRenderer = new PresentationRenderer();
                            //Convert PowerPoint slide to image as stream.
                            Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                            //Reset the stream position
                            stream.Position = 0;

                            //Download image in the browser.
                            return File(stream, "application/jpeg", "PPTXtoImage.Jpeg");                                                   
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = ex.ToString();
                    }
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose PowerPoint document to convert to image");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a PowerPoint document and then click the button to convert as a image");
            }
            return View("Index");
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