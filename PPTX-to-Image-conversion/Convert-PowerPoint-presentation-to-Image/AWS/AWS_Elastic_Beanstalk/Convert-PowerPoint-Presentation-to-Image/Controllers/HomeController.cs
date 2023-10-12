using Convert_PowerPoint_Presentation_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
using System.Reflection.Metadata;
using Syncfusion.Drawing;

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
        public IActionResult ConvertPPTXToImage()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath("wwwroot/Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PPTX document.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {                   
                    //Hooks the font substitution event.
                    pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    MemoryStream stream = new MemoryStream();
                    stream = (MemoryStream)pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                    //Unhooks the font substitution event after converting to image file.
                    pptxDoc.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                    stream.Position = 0;
                    //Download image file in the browser.
                    return File(stream, "image/jpeg", "PPTXToimage_Page1.jpeg");
                }
            }
        }

        //Sets the alternate font when a specified font is not installed in the production environment.
        private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath("wwwroot/Fonts/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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