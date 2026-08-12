using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieve the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Remove the first paragraph from the textbody of the shape.
shape.TextBody.Paragraphs.Remove(paragraph);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));