using Convert_PowerPoint_Presentation_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;


namespace Convert_PowerPoint_Presentation_to_PDF.Controllers
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

        public ActionResult ConvertPPTXtoPDF(string button)
        {
            if (button == null)
                return View("Index");
            if (Request.Form.Files != null)
            {
                if (Request.Form.Files.Count == 0)
                {
                    ViewBag.Message = string.Format("Browse a PowerPoint Presentation and then click the button to convert as a PDF document");
                    return View("Index");
                }
                // Gets the extension from file.
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();
                // Compares extension with supported extensions.
                if (extension == ".pptx")
                {
                    MemoryStream stream = new MemoryStream();
                    Request.Form.Files[0].CopyTo(stream);
                    try
                    {
                        //Open the existing PowerPoint presentation with loaded stream.
                        using (IPresentation pptxDoc = Presentation.Open(stream))
                        {
                            //Convert the PowerPoint document to PDF document.
                            using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                            {
                                //Create the MemoryStream to save the converted PDF.      
                                MemoryStream pdfStream = new MemoryStream();
                                //Save the converted PDF document to MemoryStream.
                                pdfDocument.Save(pdfStream);
                                pdfStream.Position = 0;
                                //Download PDF document in the browser.
                                return File(pdfStream, "application/pdf", "Sample.pdf");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = ex.ToString();
                    }
                }
                else
                {
                    ViewBag.Message = string.Format("Please choose PowerPoint document to convert to PDF");
                }
            }
            else
            {
                ViewBag.Message = string.Format("Browse a PowerPoint document and then click the button to convert as a PDF document");
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