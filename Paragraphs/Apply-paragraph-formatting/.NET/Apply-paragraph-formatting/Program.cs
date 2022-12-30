using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the shape in slide.
IShape textboxShape = slide.Shapes[0] as IShape;
//Get instance of a paragraph in a textbox.
IParagraph paragraph = textboxShape.TextBody.Paragraphs[0];
//Apply the first line indent of the paragraph.
paragraph.FirstLineIndent = 10;
//Apply the horizontal alignment of the paragraph to center.
paragraph.HorizontalAlignment = HorizontalAlignmentType.Left;
//Apply the left indent of the paragraph.
paragraph.LeftIndent = 8;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);