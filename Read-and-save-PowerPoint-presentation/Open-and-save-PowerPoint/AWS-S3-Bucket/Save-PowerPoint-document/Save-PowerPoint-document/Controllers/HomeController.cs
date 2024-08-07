using Amazon.S3.Transfer;
using Amazon.S3;
using Microsoft.AspNetCore.Mvc;
using Save_PowerPoint_document.Models;
using Syncfusion.Presentation;
using System.Diagnostics;

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

            //Upload the document to AWS S3
            await UploadDocumentToS3(stream);

            return Ok("PowerPoint uploaded to AWS S3 Storage.");
        }
        /// <summary>
        /// Upload file to AWS S3 cloud storage
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public async Task<MemoryStream> UploadDocumentToS3(MemoryStream stream)
        {
            //Your AWS Storage Account bucket name 
            string bucketName = "your-bucket-name";

            //Name of the PowerPoint file you want to upload
            string key = "CreatePowerPoint.pptx";

            //Configure AWS credentials and region
            var region = Amazon.RegionEndpoint.USEast1;
            var credentials = new Amazon.Runtime.BasicAWSCredentials("your-access-key", "your-secret-key");
            var config = new AmazonS3Config
            {
                RegionEndpoint = region
            };

            using (var client = new AmazonS3Client(credentials, config))
            {
                var fileTransferUtility = new TransferUtility(client);

                try
                {
                    //Upload the stream to AWS S3
                    await fileTransferUtility.UploadAsync(stream, bucketName, key);
                    Console.WriteLine("Upload completed successfully");
                }
                catch (AmazonS3Exception e)
                {
                    Console.WriteLine("Error encountered on server. Message:'{0}' when writing an object", e.Message);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unknown encountered on server. Message:'{0}' when writing an object", e.Message);
                }
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
