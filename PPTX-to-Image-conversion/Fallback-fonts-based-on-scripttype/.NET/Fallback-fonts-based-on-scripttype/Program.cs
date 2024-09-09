using Syncfusion.Office;
using Syncfusion.Presentation;
using Syncfusion.Office;
using Syncfusion.PresentationRenderer;
namespace Fallback_fonts_based_on_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Load the PowerPoint presentation into stream.
            using FileStream fileStreamInput = new FileStream(@"Data/Template.pptx", FileMode.Open, FileAccess.Read);
            //Open the existing PowerPoint presentation with loaded stream.
            using IPresentation pptxDoc = Presentation.Open(fileStreamInput);
            //Adds fallback font for "Arabic" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
            //Adds fallback font for "Hebrew" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
            //Adds fallback font for "Hindi" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Hindi, "Mangal, Nirmala UI");
            //Adds fallback font for "Chinese" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "DengXian, MingLiU");
            //Adds fallback font for "Japanese" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Mincho, MS Mincho");
            //Adds fallback font for "Thai" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
            //Adds fallback font for "Korean" script type.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");
            //Initialize the PresentationRenderer to perform image conversion.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Convert PowerPoint slide to image as stream.
            using Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
            //Reset the stream position.
            stream.Position = 0;
            //Create the output image file stream.
            using FileStream fileStreamOutput = File.Create(@"Output/Output.jpg");
            //Copy the converted image stream into created output stream.
            stream.CopyTo(fileStreamOutput);
        }
    }
}
