using Convert_PPTX_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.Diagnostics;

namespace Convert_PPTX_to_PDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public ActionResult ConvertPPTXToPDF()
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                using (FileStream inputStream = new FileStream(Path.GetFullPath("Data/Input.pptx"), FileMode.Open, FileAccess.Read))
                {
                    //Opens the existing PPTX document.
                    using (IPresentation pptxDoc = Presentation.Open(inputStream))
                    {
                        //Converts PPTX document into PDF document.
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                        {
                            //Saves the PDF document to MemoryStream.
                            pdfDocument.Save(stream);
                            stream.Position = 0;
                        }
                    }
                }
                //Download Word document in the browser.
                return File(stream, "application/pdf", "PPTXtoPDF.pdf");
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