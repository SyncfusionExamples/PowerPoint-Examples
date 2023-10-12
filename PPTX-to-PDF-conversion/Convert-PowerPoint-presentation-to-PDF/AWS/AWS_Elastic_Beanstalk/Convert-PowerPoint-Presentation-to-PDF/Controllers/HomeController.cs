using Convert_PowerPoint_Presentation_to_PDF.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.IO;
using Syncfusion.Drawing;

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
        public ActionResult ConvertPPTXtoPDF()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath("wwwroot/Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PPTX document.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Hooks the font substitution event.
                    pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Converts PPTX document into PDF document.
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        //Saves the PDF document to MemoryStream.
                        MemoryStream stream = new MemoryStream();
                        pdfDocument.Save(stream);
                        //Unhooks the font substitution event after converting to PDF.
                        pptxDoc.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                        stream.Position = 0;
                        //Download PDF document in the browser.
                        return File(stream, "application/pdf", "Sample.pdf");
                    }
                }
            }
        }
        private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
            {
                args.AlternateFontStream = new FileStream(Path.GetFullPath("wwwroot/Fonts/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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