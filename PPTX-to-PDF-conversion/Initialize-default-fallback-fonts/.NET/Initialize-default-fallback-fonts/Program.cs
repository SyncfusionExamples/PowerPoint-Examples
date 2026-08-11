using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Initialize the default FallbackFont collection.
    pptxDoc.FontSettings.FallbackFonts.InitializeDefault();
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(Path.GetFullPath(@"Output/PPTXToPDF.pdf"));
    }
}