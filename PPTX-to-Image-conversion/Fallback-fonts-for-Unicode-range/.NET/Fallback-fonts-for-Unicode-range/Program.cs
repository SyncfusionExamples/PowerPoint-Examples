using Syncfusion.Office;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Adds fallback font for specific unicode range.
// Arabic.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0600, 0x06ff, "Arial"));
// Hebrew.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Arial"));
// Hindi.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x0900, 0x097F, "Mangal"));
// Chinese.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x4E00, 0x9FFF, "DengXian"));
// Japanese.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0x3040, 0x309F, "MS Mincho"));
// Korean.
pptxDoc.FontSettings.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));
//Initialize the PresentationRenderer to perform image conversion.
pptxDoc.PresentationRenderer = new PresentationRenderer();
//Convert PowerPoint slide to image as stream.
Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
//Reset the stream position.
stream.Position = 0;
//Create the output image file stream.
using FileStream fileStreamOutput = File.Create("../../../Output.jpg");
//Copy the converted image stream into created output stream.
stream.CopyTo(fileStreamOutput);