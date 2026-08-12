using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load the PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Initialize the SubstituteFont event to set the replacement font.
    pptxDoc.FontSettings.SubstituteFont += SubstituteFont;
    //Convert the PowerPoint presentation to PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
    {
        //Save the generated PDF to the file system.
        pdfDocument.Save(Path.GetFullPath(@"Output/PPTXToPDF.pdf"));
        //Release all resources.
        pdfDocument.Close(true);
    }
}
/// <summary>
/// Sets the alternate font when a specified font is unavailable in the production environment.
/// </summary>
/// <param name="sender">FontSettings type of the Presentation in which the specified font is used but unavailable in production environment. </param>
/// <param name="args">Retrieves the unavailable font name and receives the substitute font name for conversion. </param>
static void SubstituteFont(object sender, SubstituteFontEventArgs args)
{
    if (args.OriginalFontName == "Arial Unicode MS")
        args.AlternateFontName = "Arial";
    else
        args.AlternateFontName = "Times New Roman";
}