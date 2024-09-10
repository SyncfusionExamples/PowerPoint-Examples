using Syncfusion.Office;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the PresentationRenderer to perform image conversion.
pptxDoc.PresentationRenderer = new PresentationRenderer();
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
//Convert PowerPoint slide to image as stream.
Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
//Reset the stream position.
stream.Position = 0;
//Create the output image file stream.
using FileStream fileStreamOutput = File.Create(Path.GetFullPath("Output/Output.jpg"));
//Copy the converted image stream into created output stream.
stream.CopyTo(fileStreamOutput);