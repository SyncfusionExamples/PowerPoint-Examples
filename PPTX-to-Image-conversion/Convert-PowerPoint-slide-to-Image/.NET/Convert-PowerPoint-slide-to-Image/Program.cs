using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the PresentationRenderer to perform image conversion.
pptxDoc.PresentationRenderer = new PresentationRenderer();
//Convert PowerPoint slide to image as stream.
using Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
//Create the output image file stream.
using FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output.jpg"));
//Copy the converted image stream into created output stream.
stream.CopyTo(fileStreamOutput);