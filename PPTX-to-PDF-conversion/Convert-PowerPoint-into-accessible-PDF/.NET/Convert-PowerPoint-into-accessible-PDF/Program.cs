using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
{
    //Instantiate PresentationToPdfConverterSettings.
    PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
    //Enable a flag to preserve structured document tags in the converted PDF document.
    pdfConverterSettings.AutoTag = true;
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(@"Output/PPTXToPDF.pdf");
    }
}