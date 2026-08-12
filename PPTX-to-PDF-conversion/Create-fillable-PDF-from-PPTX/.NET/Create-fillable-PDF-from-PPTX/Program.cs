using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Instantiate PresentationToPdfConverterSettings.
    PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
    //Enable a flag to preserve form fields by converting shapes with names starting with 'FormField_' into editable text form fields in the PDF.
    settings.PreserveFormFields = true;
    //Convert the PowerPoint presentation to a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the PDF document to the file system.
        pdfDocument.Save(Path.GetFullPath(@"Output/PPTXToPDF.pdf"));
    }
}