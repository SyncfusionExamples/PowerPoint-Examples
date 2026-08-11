using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Fallback_symbols_based_on_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                //Adds fallback font for basic symbols like bullet characters.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Symbols, "Segoe UI Symbol, Arial Unicode MS, Wingdings");
                //Adds fallback font for mathematics symbols.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Mathematics, "Cambria Math, Noto Sans Math, Segoe UI Symbol, Arial Unicode MS");
                //Adds fallback font for emojis.
                pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Emoji, "Segoe UI Emoji, Noto Color Emoji, Arial Unicode MS");
                //Convert the PowerPoint document to PDF document.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Save the PDF document to the file system.
                    pdfDocument.Save(@"../../../Output/PPTXToPDF.pdf");
                }
            }
        }
    }
}
