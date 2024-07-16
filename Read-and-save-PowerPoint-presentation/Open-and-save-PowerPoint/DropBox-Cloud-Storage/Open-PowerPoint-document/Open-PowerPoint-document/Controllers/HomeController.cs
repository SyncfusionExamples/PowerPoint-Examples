using Dropbox.Api;
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
            try
            {
                //Retrieve the document from DropBox
                MemoryStream stream = await GetDocumentFromDropBox();

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
        /// Download file from DropBox cloud storage
        /// </summary>
        /// <returns></returns>
        public async Task<MemoryStream> GetDocumentFromDropBox()
        {
            //Define the access token for authentication with the Dropbox API
            var accessToken = "Access_Token";

            //Define the file path in Dropbox where the file is located
            var filePathInDropbox = "FilePath";

            try
            {
                //Create a new DropboxClient instance using the provided access token
                using (var dbx = new DropboxClient(accessToken))
                {
                    //Start a download request for the specified file in Dropbox
                    using (var response = await dbx.Files.DownloadAsync(filePathInDropbox))
                    {
                        //Get the content of the downloaded file as a stream
                        var content = await response.GetContentAsStreamAsync();

                        MemoryStream stream = new MemoryStream();
                        content.CopyTo(stream);
                        return stream;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from DropBox: {ex.Message}");
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
