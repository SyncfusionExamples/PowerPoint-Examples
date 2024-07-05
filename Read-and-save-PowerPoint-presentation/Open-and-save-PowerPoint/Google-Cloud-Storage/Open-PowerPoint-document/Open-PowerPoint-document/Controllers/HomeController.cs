using Google.Apis.Auth.OAuth2;
using Google.Cloud.Storage.V1;
using Microsoft.AspNetCore.Mvc;
using Open_PowerPoint_document.Models;
using Syncfusion.Presentation;
using System.Diagnostics;
using System.IO;

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
            //Download the file from Google
            MemoryStream memoryStream = await GetDocumentFromGoogle();

            //Create an instance of PowerPoint Presentation file
            using (IPresentation pptxDocument = Presentation.Open(memoryStream))
            {
                //Get the first slide from the PowerPoint presentation
                ISlide slide = pptxDocument.Slides[0];

                //Get the first shape of the slide
                IShape shape = slide.Shapes[0] as IShape;

                //Change the text of the shape
                if (shape.TextBody.Text == "Company History")
                    shape.TextBody.Text = "Company Profile";

                //Saving the PowerPoint to a MemoryStream 
                MemoryStream outputStream = new MemoryStream();
                pptxDocument.Save(outputStream);

                //Download the PowerPoint file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/powerpoint");
                fileStreamResult.FileDownloadName = "EditPowerPoint.pptx";
                return fileStreamResult;
            }
        }
        /// <summary>
        /// Download file from Google
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromGoogle()
        {
            try
            {
                //Your bucket name
                string bucketName = "Your_bucket_name";

                //Your service account key file path
                string keyPath = "credentials.json";

                //Name of the file to download from the Google Cloud Storage
                string fileName = "PowerPointTemplate.pptx";

                //Create Google Credential from the service account key file
                GoogleCredential credential = GoogleCredential.FromFile(keyPath);

                //Instantiates a storage client to interact with Google Cloud Storage
                StorageClient storageClient = StorageClient.Create(credential);

                //Download a file from Google Cloud Storage
                MemoryStream memoryStream = new MemoryStream();
                await storageClient.DownloadObjectAsync(bucketName, fileName, memoryStream);
                memoryStream.Position = 0;

                return memoryStream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from Google Cloud Storage: {ex.Message}");
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
