using Syncfusion.Drawing;
using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Use a sets of default FallbackFont collection to IPresentation.
pptxDoc.FontSettings.FallbackFonts.InitializeDefault();
// Customize a default fallback font name.
FallbackFonts fallbackFonts = pptxDoc.FontSettings.FallbackFonts;
foreach (FallbackFont fallbackFont in fallbackFonts) 
{
    //Customize a default fallback font name as "David" for the Hebrew script.
    if (fallbackFont.ScriptType == ScriptType.Hebrew)
        fallbackFont.FontNames = "David";
}
//Create the MemoryStream to save the converted PDF.
using MemoryStream pdfStream = new();
//Convert the PowerPoint document to PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
//Save the converted PDF document to MemoryStream.
pdfDocument.Save(pdfStream);
pdfStream.Position = 0;
//Create the output PDF file stream.
using FileStream fileStreamOutput = File.Create("../../../PPTXToPDF.pdf");
//Copy the converted PDF stream into created output PDF stream.
pdfStream.CopyTo(fileStreamOutput);