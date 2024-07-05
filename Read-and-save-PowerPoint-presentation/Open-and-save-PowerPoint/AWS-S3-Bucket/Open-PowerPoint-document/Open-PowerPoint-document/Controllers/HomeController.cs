using Amazon.S3;
using Microsoft.AspNetCore.Mvc;
using Open_PowerPoint_document.Models;
using Syncfusion.Presentation;
using System.Diagnostics;

namespace Open_PowerPoint_document.Controllers
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
        public async Task<IActionResult> EditDocument()
        {
            try
            {
                //Retrieve the document from AWS S3
                MemoryStream stream = await GetDocumentFromS3();

                //Set the position to the beginning of the MemoryStream
                stream.Position = 0;

                //Create an instance of PowerPoint Presentation file
                using (IPresentation pptxDocument = Presentation.Open(stream))
                {
                    //Get the first slide from the PowerPoint presentation
                    ISlide slide = pptxDocument.Slides[0];

                    //Get the first shape of the slide
                    IShape shape = slide.Shapes[0] as IShape;

                    //Change the text of the shape
                    if (shape.TextBody.Text == "Company History")
                        shape.TextBody.Text = "Company Profile";

                    //Saving the PowerPoint file to a MemoryStream 
                    MemoryStream outputStream = new MemoryStream();
                    pptxDocument.Save(outputStream);

                    //Download the PowerPoint file in the browser
                    FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/powerpoint");
                    fileStreamResult.FileDownloadName = "EditPowerPoint.pptx";
                    return fileStreamResult;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return Content("Error occurred while processing the file.");
            }
        }
        /// <summary>
        /// Download file from AWS S3 cloud storage
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromS3()
        {
            //Your AWS Storage Account bucket name 
            string bucketName = "your-bucket-name";

            //Name of the PowerPoint file you want to load from AWS S3
            string key = "PowerPointTemplate.pptx";

            //Configure AWS credentials and region
            var region = Amazon.RegionEndpoint.USEast1;
            var credentials = new Amazon.Runtime.BasicAWSCredentials("your-access-key", "your-secret-key");
            var config = new AmazonS3Config
            {
                RegionEndpoint = region
            };

            try
            {
                using (var client = new AmazonS3Client(credentials, config))
                {
                    //Create a MemoryStream to copy the file content
                    MemoryStream stream = new MemoryStream();

                    //Download the file from S3 into the MemoryStream
                    var response = await client.GetObjectAsync(new Amazon.S3.Model.GetObjectRequest
                    {
                        BucketName = bucketName,
                        Key = key
                    });

                    //Copy the response stream to the MemoryStream
                    await response.ResponseStream.CopyToAsync(stream);

                    return stream;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from S3: {ex.Message}");
                throw; // or handle the exception as needed
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
