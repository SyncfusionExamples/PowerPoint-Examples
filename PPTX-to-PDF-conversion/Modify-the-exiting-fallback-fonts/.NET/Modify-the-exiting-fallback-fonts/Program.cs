using Syncfusion.Drawing;
using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
{
    //Initialize the default FallbackFont collection.
    pptxDoc.FontSettings.FallbackFonts.InitializeDefault();
    //Customize a default fallback font name.
    FallbackFonts fallbackFonts = pptxDoc.FontSettings.FallbackFonts;
    foreach (FallbackFont fallbackFont in fallbackFonts)
    {
        //Customize the default fallback font name to "David" for the Hebrew script.
        if (fallbackFont.ScriptType == ScriptType.Hebrew)
            fallbackFont.FontNames = "David";
    }
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(@"Output/PPTXToPDF.pdf");
    }
}