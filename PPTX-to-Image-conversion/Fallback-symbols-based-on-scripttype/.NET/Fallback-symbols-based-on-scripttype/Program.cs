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
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
            {
                //Adds fallback font for basic symbols like bullet characters.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Symbols, "Segoe UI Symbol, Arial Unicode MS, Wingdings");
                //Adds fallback font for mathematics symbols.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Mathematics, "Cambria Math, Noto Sans Math, Segoe UI Symbol, Arial Unicode MS");
                //Adds fallback font for emojis.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Emoji, "Segoe UI Emoji, Noto Color Emoji, Arial Unicode MS");
                //Initialize the PresentationRenderer to perform image conversion.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert PowerPoint slide to image as stream.
                using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                {
                    //Reset the stream position.
                    stream.Position = 0;
                    //Create the output image file stream.
                    using (FileStream fileStreamOutput = File.Create(@"../../../Output/Output.jpg"))
                    {
                        //Copy the converted image stream into created output stream.
                        stream.CopyTo(fileStreamOutput);
                    }
                }
            }
        }
    }
}
