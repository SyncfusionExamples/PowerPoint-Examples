using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Load the PowerPoint presentation into a stream.
using FileStream fileStreamInput = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open the existing PowerPoint presentation with the loaded stream.
using IPresentation pptxDoc = Presentation.Open(fileStreamInput) ;
//Instantiation of the PresentationToPdfConverterSettings.
PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
//Enable a flag to preserve structured document tags in the converted PDF document.               
pdfConverterSettings.AutoTag = true;
//Convert the PowerPoint document to a PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings);
//Create new instance of file stream.
using FileStream fileStreamOutput = new(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create);
//Save the generated PDF to file stream.
pdfDocument.Save(fileStreamOutput);