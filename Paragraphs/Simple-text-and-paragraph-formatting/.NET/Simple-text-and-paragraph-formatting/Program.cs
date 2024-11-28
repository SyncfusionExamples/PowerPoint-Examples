using Syncfusion.Presentation;

//Load PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieve the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Set center alignment for the paragraph.
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Retrieve the first TextPart of the shape.
ITextPart textPart = paragraph.TextParts[0];
//Set bold for the text.
textPart.Font.Bold = true;
//Save the PowerPoint Presentation.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
