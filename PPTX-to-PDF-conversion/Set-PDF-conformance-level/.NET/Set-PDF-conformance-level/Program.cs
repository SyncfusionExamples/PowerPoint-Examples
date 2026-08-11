using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Initialize the conversion settings.
    PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
    //Set the PDF conformance level to A1B.
    pdfConverterSettings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDoc = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings))
    {
        //Save the PDF document to the file system.
        pdfDoc.Save(Path.GetFullPath(@"Output/PPTXToPDF.pdf"));
    }
}