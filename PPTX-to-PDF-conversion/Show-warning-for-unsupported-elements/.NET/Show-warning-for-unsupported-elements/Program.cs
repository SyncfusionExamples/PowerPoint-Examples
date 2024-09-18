using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Instantiation of PresentationToPdfConverterSettings.
PresentationToPdfConverterSettings settings = new()
{
    //Get all the warnings into the collection.
    Warning = new DocumentWarning()
};
//Convert the PowerPoint Presentation into PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, settings);
//If the PowerPoint to Pdf conversion has been stopped, IsCanceled value will be True, otherwise false.
if (!PresentationToPdfConverter.IsCanceled)
{
    //Save the PDF file.
    using FileStream outputStream = new(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
    pdfDocument.Save(outputStream);
    outputStream.Position = 0;
}
else
{
    Console.WriteLine("PowerPoint to PDF conversion is stopped , please press any key to exit the application");
    Console.ReadKey();
}

/// <summary>
/// DocumentWarning class implements the IWarning interface.
/// </summary>
/// <seealso cref="IWarning" />
public class DocumentWarning : IWarning
{
    /// <summary>
    /// Get the Boolen value whether to continue conversion or not.
    /// </summary>
    /// <param name="warningInfo">Collection of warnings</param>
    /// <returns></returns>
    public bool ShowWarnings(List<WarningInfo> warningInfo)
    {
        //By default to perform the PowerPoint to PDF conversion by setting the isContinueConversion as true.
        bool isContinueConversion = true;
        foreach (WarningInfo warning in warningInfo)
        {
            //Since there are warnings in the PowerPoint presentation the value of isContinueConversion will be set as false.
            isContinueConversion = false;
            //Print the description of the Warining
            Console.WriteLine(warning.Description);
            if (warning.Description.Contains("Metafile") || warning.Description.Contains("Chart"))
            {
                Console.WriteLine("Type [Y] if you want Do you want to continue Presentation to Pdf conversion or Type [N] to cancel the conversion");
                String confrimation = "y";
                //Based on warning.WarningType enumeration, you can do your manipulation.
                //Skips the PowerPoint to Pdf conversion by setting isContinueConversion value as false.
                //Continue the PowerPoint to PDF conversion by setting the isContinueConversion as true.
                if (confrimation.ToLower().Equals("y"))
                    isContinueConversion = true;
                else
                    isContinueConversion = false;
            }
        }
        return isContinueConversion;
    }
}