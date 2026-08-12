using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
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
clonedPresentation.Save(Path.GetFullPath(@"Output/Result.pptx"));