using Amazon.Lambda.Core;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
using Syncfusion.Drawing;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Convert_PowerPoint_Presentation_to_PDF;

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
        //Open the file as Stream
        using (FileStream fileStreamInput = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            //Open the existing PowerPoint presentation with loaded stream.
            using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
            {
                //Hooks the font substitution event.
                pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                //Convert the PowerPoint document to PDF document.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Create the MemoryStream to save the converted PDF.      
                    MemoryStream pdfStream = new MemoryStream();
                    //Save the converted PDF document to MemoryStream.
                    pdfDocument.Save(pdfStream);
                    //Unhooks the font substitution event after converting to PDF.
                    pptxDoc.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                    pdfStream.Position = 0;
                    return Convert.ToBase64String(pdfStream.ToArray());
                }
            }           
        }
    }

    //Set the alternate font when a specified font is not installed in the production environment.
    private void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
    {
        if (args.OriginalFontName == "Calibri" && args.FontStyle == FontStyle.Regular)
            args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/calibri.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        else
            args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/times.ttf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }
}
