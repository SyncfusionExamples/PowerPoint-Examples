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
            using FileStream fileStreamInput = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read);
            //Open the existing PowerPoint presentation with loaded stream.
            using IPresentation pptxDoc = Presentation.Open(fileStreamInput);
            //Adds fallback font for basic symbols like bullet characters.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Symbols, "Segoe UI Symbol, Arial Unicode MS, Wingdings");
            //Adds fallback font for mathematics symbols.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Mathematics, "Cambria Math, Noto Sans Math, Segoe UI Symbol, Arial Unicode MS");
            //Adds fallback font for emojis.
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Emoji, "Segoe UI Emoji, Noto Color Emoji, Arial Unicode MS");
            //Initialize the PresentationRenderer to perform image conversion.
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Convert PowerPoint slide to image as stream.
            using Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
            //Reset the stream position.
            stream.Position = 0;
            //Create the output image file stream.
            using FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Output/Output.jpg"));
            //Copy the converted image stream into created output stream.
            stream.CopyTo(fileStreamOutput);
        }
    }
}
