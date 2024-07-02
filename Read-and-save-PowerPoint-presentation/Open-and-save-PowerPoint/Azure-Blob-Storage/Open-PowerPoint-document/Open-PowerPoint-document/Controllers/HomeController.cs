using Azure.Storage.Blobs.Models;
using Azure.Storage.Blobs;
using Microsoft.AspNetCore.Mvc;
using Open_PowerPoint_document.Models;
using System.Diagnostics;
using Syncfusion.Presentation;

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
            //Your Azure Storage Account connection string
            string connectionString = "Your_connection_string";

            //Name of the Azure Blob Storage container
            string containerName = "Your_container_name";

            //Name of the Powerpoint file you want to load
            string blobName = "PowerPointTemplate.pptx";

            try
            {
                //Retrieve the document from Azure
                MemoryStream stream = await GetDocumentFromAzure(connectionString, containerName, blobName);

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

                    //Saving the PowerPoint document to a MemoryStream 
                    MemoryStream outputStream = new MemoryStream();
                    pptxDocument.Save(outputStream);

                    //Set the position as '0'
                    outputStream.Position = 0;

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
        public async Task<MemoryStream> GetDocumentFromAzure(string connectionString, string containerName, string blobName)
        {
            try
            {
                //Download the PowerPoint document from Azure Blob Storage
                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
                BlobClient blobClient = containerClient.GetBlobClient(blobName);
                BlobDownloadInfo download = await blobClient.DownloadAsync();

                //Create a MemoryStream to copy the file content
                MemoryStream stream = new MemoryStream();
                await download.Content.CopyToAsync(stream);

                return stream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from Azure Blob Storage: {ex.Message}");
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
