using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieve the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Retrieve the first TextPart of the shape.
ITextPart textPart = paragraph.TextParts[0];
//Modify the text content of the TextPart.
textPart.Text = "Hello Presentation";
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));