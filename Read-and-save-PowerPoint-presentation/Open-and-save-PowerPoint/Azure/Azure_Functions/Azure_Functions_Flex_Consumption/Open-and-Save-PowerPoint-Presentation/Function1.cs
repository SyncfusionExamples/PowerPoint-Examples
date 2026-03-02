using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Syncfusion.Presentation;

namespace Open_and_Save_PowerPoint_Presentation;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("OpenandSavePowerPointPresentation")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        try
        {
            // Create a memory stream to hold the incoming request body (PowerPoint Presentation bytes)
            await using MemoryStream inputStream = new MemoryStream();
            // Copy the request body into the memory stream
            await req.Body.CopyToAsync(inputStream);
            // Check if the stream is empty (no file content received)
            if (inputStream.Length == 0)
                return new BadRequestObjectResult("No file content received in request body.");
            // Reset stream position to the beginning for reading
            inputStream.Position = 0;
            // Load the PowerPoint Presentation from the stream
            using IPresentation pptxDoc = Presentation.Open(inputStream);
            //Gets the first slide from the PowerPoint presentation
            ISlide slide = pptxDoc.Slides[0];
            //Gets the first shape of the slide
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            MemoryStream memoryStream = new MemoryStream();
            //Saves the PowerPoint document file.
            pptxDoc.Save(memoryStream);
            memoryStream.Position = 0;
            var bytes = memoryStream.ToArray();
            return new FileContentResult(bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            {
                FileDownloadName = "presentation.pptx"
            };
        }
        catch(Exception ex)
        {
            // Log the error with details for troubleshooting
            _logger.LogError(ex, "Error open and save PowerPoint Presentation.");
            // Prepare error message including exception details
            var msg = $"Exception: {ex.Message}\n\n{ex}";
            // Return a 500 Internal Server Error response with the message
            return new ContentResult { StatusCode = 500, Content = msg, ContentType = "text/plain; charset=utf-8" };
        }
    }
}