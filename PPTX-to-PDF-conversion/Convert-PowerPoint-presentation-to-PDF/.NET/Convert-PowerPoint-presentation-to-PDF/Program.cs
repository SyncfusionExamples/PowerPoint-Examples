using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath("Data/Template.pptx")))
{
    //Convert the PowerPoint presentation to PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(@"Output\PPTXToPDF.pdf");
    }
}
