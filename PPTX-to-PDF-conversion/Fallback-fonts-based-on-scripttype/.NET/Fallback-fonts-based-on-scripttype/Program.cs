using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
namespace Fallback_fonts_based_on_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Load the PowerPoint presentation into stream
            using FileStream fileStreamInput = new FileStream(@"../../../Data/Template.pptx", FileMode.Open, FileAccess.Read);
            //Open the existing PowerPoint presentation with loaded stream
            using IPresentation pptxDoc = Presentation.Open(fileStreamInput);
            //Adds fallback font for "Arabic" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
            //Adds fallback font for "Hebrew" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
            //Adds fallback font for "Hindi" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Hindi, "Mangal, Nirmala UI");
            //Adds fallback font for "Chinese" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "DengXian, MingLiU");
            //Adds fallback font for "Japanese" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Mincho, MS Mincho");
            //Adds fallback font for "Thai" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
            //Adds fallback font for "Korean" script type
            pptxDoc.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");
            //Create the MemoryStream to save the converted PDF
            using MemoryStream pdfStream = new MemoryStream();
            //Convert the PowerPoint document to PDF document
            using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
            //Save the converted PDF document to MemoryStream
            pdfDocument.Save(pdfStream);
            pdfStream.Position = 0;
            //Create the output PDF file stream
            using FileStream fileStreamOutput = File.Create(@"../../../PPTXToPDF.pdf");
            //Copy the converted PDF stream into created output PDF stream
            pdfStream.CopyTo(fileStreamOutput);
        }
    }
}
