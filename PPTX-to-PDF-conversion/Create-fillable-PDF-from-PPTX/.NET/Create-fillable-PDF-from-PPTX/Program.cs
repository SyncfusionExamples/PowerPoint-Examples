using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the PowerPoint file stream. 
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load an existing PowerPoint Presentation. 
    using (IPresentation pptxDoc = Presentation.Open(fileStream))
    {
        // Create new instance for PresentationToPdfConverterSettings
        PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
        //Enables a flag to preserve form fields by converting shapes with names starting with 'FormField_' into editable text form fields in the PDF.
        settings.PreserveFormFields = true;
        //Convert PowerPoint into PDF document. 
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
        {
            //Save the PDF file to file system. 
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}