using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
{
    //Instantiate PresentationToPdfConverterSettings.
    PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
    //Subscribe to the warnings collection.
    settings.Warning = new DocumentWarning();
    //Convert the PowerPoint presentation into a PDF document.
    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, settings))
    {
        //If the PowerPoint to PDF conversion has been stopped, IsCanceled is true; otherwise it is false.
        if (!PresentationToPdfConverter.IsCanceled)
        {
            //Save the PDF file.
            pdfDocument.Save(@"Output/PPTXToPDF.pdf");
        }
        else
        {
            Console.WriteLine("PowerPoint to PDF conversion is stopped. Press any key to exit the application.");
            Console.ReadKey();
        }
    }
}

/// <summary>
/// DocumentWarning class implements the IWarning interface.
/// </summary>
/// <seealso cref="IWarning" />
public class DocumentWarning : IWarning
{
    /// <summary>
    /// Gets the Boolean value indicating whether to continue the conversion.
    /// </summary>
    /// <param name="warningInfo">Collection of warnings</param>
    /// <returns>True to continue the conversion; otherwise, false.</returns>
    public bool ShowWarnings(List<WarningInfo> warningInfo)
    {
        //By default, continue the PowerPoint to PDF conversion by setting isContinueConversion to true.
        bool isContinueConversion = true;
        foreach (WarningInfo warning in warningInfo)
        {
            //Mark the conversion as interrupted when warnings are present; the loop below can override this.
            isContinueConversion = false;
            //Print the description of the warning.
            Console.WriteLine(warning.Description);
            if (warning.Description.Contains("Metafile") || warning.Description.Contains("Chart"))
            {
                Console.WriteLine("Type [Y] if you want to continue the Presentation to PDF conversion, or [N] to cancel the conversion.");
                String confirmation = Console.ReadLine();
                //Based on the WarningType enumeration, you can perform custom logic here.
                //Continue the PowerPoint to PDF conversion when the user enters Y; otherwise, cancel it.
                if (confirmation.ToLower().Equals("y"))
                    isContinueConversion = true;
                else
                    isContinueConversion = false;
            }
        }
        return isContinueConversion;
    }
}