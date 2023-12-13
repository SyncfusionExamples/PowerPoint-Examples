using Syncfusion.Drawing;
using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Add custom fallback font names.
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
//Create the MemoryStream to save the converted PDF.
using MemoryStream pdfStream = new MemoryStream();
//Convert the PowerPoint document to PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
//Save the converted PDF document to MemoryStream.
pdfDocument.Save(pdfStream);
pdfStream.Position = 0;
//Create the output PDF file stream.
using FileStream fileStreamOutput = File.Create("../../../PPTXToPDF.pdf");
//Copy the converted PDF stream into created output PDF stream.
pdfStream.CopyTo(fileStreamOutput);