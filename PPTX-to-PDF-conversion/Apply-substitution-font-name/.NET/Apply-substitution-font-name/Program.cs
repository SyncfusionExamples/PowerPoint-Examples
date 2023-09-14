using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
// Initialize the 'SubstituteFont' event to set the replacement font.
pptxDoc.FontSettings.SubstituteFont += SubstituteFont;
//Convert the PowerPoint presentation to PDF file.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
//Create new instance of file stream.
FileStream pdfStream = new("../../../PPTXToPDF.pdf", FileMode.Create);
//Save the generated PDF to file stream.
pdfDocument.Save(pdfStream);

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