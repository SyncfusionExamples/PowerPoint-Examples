using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the PresentationRenderer to perform image conversion.
pptxDoc.PresentationRenderer = new PresentationRenderer();
//Use a sets of default FallbackFont collection to IPresentation.
pptxDoc.FontSettings.InitializeFallbackFonts();
// Customize a default fallback font name.
// Modify the Hebrew script default font name as "David".
pptxDoc.FontSettings.FallbackFonts[5].FontNames = "David";
//Convert PowerPoint slide to image as stream.
Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
//Reset the stream position.
stream.Position = 0;
//Create the output image file stream.
using FileStream fileStreamOutput = File.Create("../../../Output.jpg");
//Copy the converted image stream into created output stream.
stream.CopyTo(fileStreamOutput);