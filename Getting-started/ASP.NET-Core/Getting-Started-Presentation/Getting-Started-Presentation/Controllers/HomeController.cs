using Getting_Started_Presentation.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Presentation;
using System.Diagnostics;

namespace Getting_Started_Presentation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public ActionResult CreatePresentation()
        {
            //Creates a new instance of PowerPoint Presentation.
            IPresentation pptxDoc = Presentation.Create();
            //Add a slide of Title type.
            ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Title);
            //Gets the title shape.
            IShape titleShape = slide1.Shapes[0] as IShape;
            //Removes the description shape.
            slide1.Shapes.RemoveAt(1);
            //Adds paragraph to the textbody of titleShape.
            IParagraph paragraph = titleShape.TextBody.AddParagraph("Azure Pipelines");
            //Sets the font name for the text.
            paragraph.Font.FontName = "Tahoma";
            //Applies the text formatting.
            paragraph.Font.Bold = true;

            //Add a slide of TitleAndContent type.
            ISlide slide2 = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent);
            //Remove title shape.
            slide2.Shapes.RemoveAt(0);
            //Gets the description shape.
            IShape descriptionShape = slide2.Shapes[0] as IShape;
            //Add paragraph to the textbody of description shape.
            paragraph = descriptionShape.TextBody.AddParagraph();
            //Removes the list formatting of the paragraph.
            paragraph.ListFormat.Type = ListType.None;
            //Applies font size of the paragraph
            paragraph.Font.FontSize = 22;
            //Applies the line spacing of the paragraph.
            paragraph.LineSpacing = 40;
            //Applies the left indent of the paragraph.
            paragraph.LeftIndent = 0;
            //Adds text to the Paragraph
            paragraph.Text = "To make it easier for developers to move between different parts of the software stack, " +
                "the .NET Engineering Services team worked to bring all repos under a common directory structure, set of commands, and build-and-test logic. " +
                "The team eliminated further barriers to productivity by moving all existing workflows from the different CI systems into a single system in Azure Pipelines. " +
                "Because the system is fully managed, they no longer have the burden of managing the CI server infrastructure.";


            //Add a slide of TitleAndContent type.
            ISlide slide3 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            //Add bullet points to the slide.
            IShape bulletPointsShape = slide3.AddTextBox(37, 123, 480, 215);
            //Adds a new paragraph with the text.
            IParagraph headingParagraph = bulletPointsShape.TextBody.AddParagraph("Azure Pipelines has the following advantages:");
            //Applies the line spacing of the paragraph.
            headingParagraph.LineSpacing = 40;
            //Applies font size of the paragraph.
            headingParagraph.Font.FontSize = 20;
            //Applies the text formatting.
            headingParagraph.Font.Bold = true;
            //Removes the list formatting of the paragraph.
            headingParagraph.ListFormat.Type = ListType.None;
            //Adds a new paragraph with the text.
            IParagraph bulletedParagraph = bulletPointsShape.TextBody.AddParagraph("Works with any language or platform.");
            bulletedParagraph.Font.FontSize = 20;
            bulletedParagraph.ListFormat.Type = ListType.Bulleted;
            bulletedParagraph.LineSpacing = 40;
            //Adds another paragraph with the text.
            bulletedParagraph = bulletPointsShape.TextBody.AddParagraph("Deploys to different types of targets at the same time.");
            bulletedParagraph.Font.FontSize = 20;
            bulletedParagraph.ListFormat.Type = ListType.Bulleted;
            bulletedParagraph.LineSpacing = 40;
            //Adds another paragraph with the text.
            bulletedParagraph = bulletPointsShape.TextBody.AddParagraph("Builds on Windows, Linux, or Mac machines.");
            bulletedParagraph.Font.FontSize = 20;
            bulletedParagraph.ListFormat.Type = ListType.Bulleted;
            bulletedParagraph.LineSpacing = 40;
            //Adds another paragraph with the text.
            bulletedParagraph = bulletPointsShape.TextBody.AddParagraph("Combines with GitHub.");
            bulletedParagraph.Font.FontSize = 20;
            bulletedParagraph.LineSpacing = 40;
            bulletedParagraph.ListFormat.Type = ListType.Bulleted;
            //Adds another paragraph with the text.
            bulletedParagraph = bulletPointsShape.TextBody.AddParagraph("Works with open-source projects.");
            bulletedParagraph.Font.FontSize = 20;
            bulletedParagraph.LineSpacing = 40;
            bulletedParagraph.ListFormat.Type = ListType.Bulleted;

            //Gets a picture as stream.
            Stream imageStream = new FileStream("Data/Image.png", FileMode.Open, FileAccess.Read);
            //Adds the picture to a slide by specifying its size and position.
            slide3.Shapes.AddPicture(imageStream, 520, 115, 405, 250);

            //Save the PowerPoint Presentation as stream.
            MemoryStream pptxStream = new();
            pptxDoc.Save(pptxStream);
            pptxDoc.Close();
            pptxStream.Position = 0;

            //Download Powerpoint document in the browser.
            return File(pptxStream, "application/powerpoint", "Result.pptx");
        }

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
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