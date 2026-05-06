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
            //Create the MemoryStream to save the converted PDF.
            using MemoryStream pdfStream = new MemoryStream();
            //Convert the PowerPoint document to PDF document.
            using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
            //Save the converted PDF document to MemoryStream.
            pdfDocument.Save(pdfStream);
            pdfStream.Position = 0;
            //Create the output PDF file stream.
            using FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Output/PPTXToPDF.pdf"));
            //Copy the converted PDF stream into created output PDF stream.
            pdfStream.CopyTo(fileStreamOutput);
        }
    }
}
