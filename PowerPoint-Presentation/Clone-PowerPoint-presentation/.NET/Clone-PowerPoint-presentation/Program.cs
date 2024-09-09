using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Clone the Presentation.
using IPresentation clonedPresentation = pptxDoc.Clone();
//Get the first slide from the cloned PowerPoint presentation.
ISlide firstSlide = clonedPresentation.Slides[0];
//Add a textbox in a slide by specifying its position and size.
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);
//Add a paragraph in the body of the textShape.
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Add a textPart in the paragraph.
ITextPart textPart = paragraph.AddTextPart("Essential Presentation");
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
clonedPresentation.Save(outputStream);