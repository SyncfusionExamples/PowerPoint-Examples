using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
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
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));