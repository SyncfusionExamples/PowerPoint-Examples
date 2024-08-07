using Google.Apis.Auth.OAuth2;
using Google.Cloud.Storage.V1;
using Microsoft.AspNetCore.Mvc;
using Save_PowerPoint_document.Models;
using Syncfusion.Presentation;
using System.ComponentModel;
using System.Diagnostics;
using static System.Collections.Specialized.BitVector32;

namespace Save_PowerPoint_document.Controllers
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
        public async Task<IActionResult> UploadDocument()
        {
            //Create a new instance of PowerPoint Presentation file
            IPresentation pptxDocument = Presentation.Create();

            //Add a new slide to file and apply background color
            ISlide slide = pptxDocument.Slides.Add(SlideLayoutType.TitleOnly);

            //Specify the fill type and fill color for the slide background 
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph("Company History").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Add description content to the slide by adding a new TextBox
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = "IMN Solutions PVT LTD is the software company, established in 1987, by George Milton. The company has been listed as the trusted partner for many high-profile organizations since 1988 and got awards for quality products from reputed organizations.";

            //Add bullet points to the slide
            IShape bulletPointsShape = slide.AddTextBox(53.22, 270, 437.90, 116.32);

            //Add a paragraph for a bullet point
            IParagraph firstPara = bulletPointsShape.TextBody.AddParagraph("The company acquired the MCY corporation for 20 billion dollars and became the top revenue maker for the year 2015.");

            //Format how the bullets should be displayed
            firstPara.ListFormat.Type = ListType.Bulleted;
            firstPara.LeftIndent = 35;
            firstPara.FirstLineIndent = -35;

            //Add another paragraph for the next bullet point
            IParagraph secondPara = bulletPointsShape.TextBody.AddParagraph("The company is participating in top open source projects in automation industry.");

            //Format how the bullets should be displayed
            secondPara.ListFormat.Type = ListType.Bulleted;
            secondPara.LeftIndent = 35;
            secondPara.FirstLineIndent = -35;

            //Gets a picture as stream
            FileStream pictureStream = new FileStream("Image.jpg", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Add an auto-shape to the slide
            IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);

            //Format the auto-shape color by setting the fill type and text
            stampShape.Fill.FillType = FillType.None;
            stampShape.TextBody.AddParagraph("IMN").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Saves the PowerPoint to MemoryStream
            MemoryStream stream = new MemoryStream();
            pptxDocument.Save(stream);            

            //Upload the file to Google
            await UploadDocumentToGoogle(stream);

            return Ok("PowerPoint uploaded to Google Cloud Storage.");
        }
        /// <summary>
        /// Upload file to Google
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public async Task<MemoryStream> UploadDocumentToGoogle(MemoryStream stream)
        {
            try
            {
                //Your bucket name
                string bucketName = "Your_bucket_name";

                //Your service account key file path
                string keyPath = "credentials.json";

                //Name of the file to upload to Google Cloud Storage
                string fileName = "CreatePowerPoint.pptx";

                //Create Google Credential from the service account key file
                GoogleCredential credential = GoogleCredential.FromFile(keyPath);

                //Instantiates a storage client to interact with Google Cloud Storage
                StorageClient storageClient = StorageClient.Create(credential);

                //Uploads a file to Google Cloud Storage
                storageClient.UploadObject(bucketName, fileName, null, stream);

                Console.WriteLine("Upload completed successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown encountered on server. Message:'{0}' when writing an object", e.Message);
            }
            return stream;
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
