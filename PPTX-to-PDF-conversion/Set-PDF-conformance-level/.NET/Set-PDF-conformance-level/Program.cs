using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Create a file stream to read the PowerPoint presentation file.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the conversion settings.
PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
//Set the Pdf conformance level to A1B
pdfConverterSettings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
//Convert the PowerPoint presentation to PDF file.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings);
//Create new instance of file stream.
using FileStream pdfStream = new(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create);
//Save the generated PDF to file stream.
pdfDocument.Save(pdfStream);