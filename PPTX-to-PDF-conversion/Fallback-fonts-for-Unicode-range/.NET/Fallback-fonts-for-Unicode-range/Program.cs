using Syncfusion.Drawing;
using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
{
    //Adds fallback font for specific unicode range.
    // Arabic.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0600, 0x06ff, "Arial"));
    // Hebrew.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Arial"));
    // Hindi.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0900, 0x097F, "Mangal"));
    // Chinese.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x4E00, 0x9FFF, "DengXian"));
    // Japanese.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x3040, 0x309F, "MS Mincho"));
    // Korean.
    pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(@"Output/PPTXToPDF.pdf");
    }
}