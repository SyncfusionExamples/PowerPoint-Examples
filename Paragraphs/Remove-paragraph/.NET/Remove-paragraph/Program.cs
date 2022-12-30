using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieve the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Remove the first paragraph from the textbody of the shape.
shape.TextBody.Paragraphs.Remove(paragraph);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);