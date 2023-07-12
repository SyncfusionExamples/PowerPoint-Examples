using Amazon.Lambda.Core;
using Syncfusion.Presentation;
// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Open_and_Save_PowerPoint_Presentation;

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
        //Load an existing PowerPoint file.
        string filePath = Path.GetFullPath(@"Data/Template.pptx");
        using (IPresentation pptxDoc = Presentation.Open(filePath))
        {
            //Get the first slide from the PowerPoint presentation.
            ISlide slide = pptxDoc.Slides[0];
            //Get the first shape of the slide.
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape.
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Save the PowerPoint file.
            MemoryStream stream = new MemoryStream();
            pptxDoc.Save(stream);
            return Convert.ToBase64String(stream.ToArray());            
        }       
    }
}
