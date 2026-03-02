using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using SkiaSharp;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;

namespace Convert_PowerPoint_Presentation_to_Image;

public class Function1
{
    private readonly ILogger<Function1> _logger;

    public Function1(ILogger<Function1> logger)
    {
        _logger = logger;
    }

    [Function("ConvertPowerPointtoImage")]
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
            // Attach font substitution handler to manage missing fonts
            pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
            // Initialize the PresentationRenderer to perform image conversion.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            // Convert PowerPoint slide to image as stream.
            Stream imageStream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Png);
            // Reset the stream position.
            imageStream.Position = 0;
            // Create a memory stream to hold the Image output
            await using MemoryStream outputStream = new MemoryStream();
            // Copy the contents of the image stream to the memory stream.
            await imageStream.CopyToAsync(outputStream);
            // Convert the Image stream to a byte array
            var imageBytes = outputStream.ToArray();
            //Reset the stream position.
            imageStream.Position = 0;
            // Reset stream position to the beginning for reading
            outputStream.Position = 0;
            // Create a file result to return the PNG as a downloadable file
            return new FileContentResult(imageBytes, "image/png")
            {
                FileDownloadName = "slide-1.png"
            };
        }
        catch (Exception ex)
        {
            // Log the error with details for troubleshooting
            _logger.LogError(ex, "Error converting PPTX to Image.");
            // Prepare error message including exception details
            var msg = $"Exception: {ex.Message}\n\n{ex}";
            // Return a 500 Internal Server Error response with the message
            return new ContentResult { StatusCode = 500, Content = msg, ContentType = "text/plain; charset=utf-8" };
        }
    }
    /// <summary>
    /// Event handler for font substitution during PDF conversion
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="args"></param>
    private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
    {
        // Define the path to the Fonts folder in the application base directory
        string fontsFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
        // If the original font is Calibri, substitute with calibri-regular.ttf
        if (args.OriginalFontName == "Calibri")
        {
            args.AlternateFontStream = File.OpenRead(Path.Combine(fontsFolder, "calibri-regular.ttf"));
        }
        // Otherwise, substitute with Times New Roman
        else
        {
            args.AlternateFontStream = File.OpenRead(Path.Combine(fontsFolder, "Times New Roman.ttf"));
        }
    }
}