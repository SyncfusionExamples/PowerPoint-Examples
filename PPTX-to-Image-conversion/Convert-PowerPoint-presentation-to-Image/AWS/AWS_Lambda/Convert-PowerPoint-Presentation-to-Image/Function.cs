using Amazon.Lambda.Core;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Convert_PowerPoint_Presentation_to_Image;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(string input, ILambdaContext context)
    {
        string filePath = Path.GetFullPath(@"Data/Input.pptx");
        //Open the existing PowerPoint presentation with loaded stream.
        using (IPresentation pptxDoc = Presentation.Open(filePath))
        {
            //Initialize the PresentationRenderer to perform image conversion.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Convert PowerPoint slide to image as stream.
            Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
            //Reset the stream position.
            stream.Position = 0;
            // Create a memory stream to save the image file.
            MemoryStream memoryStream = new MemoryStream();
            stream.CopyTo(memoryStream);
            return Convert.ToBase64String(memoryStream.ToArray());
        }
    }
}
