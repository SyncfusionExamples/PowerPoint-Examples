using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Retrieves the first paragraph of the shape.
IParagraph paragraph = shape.TextBody.Paragraphs[0];
//Retrieves the first TextPart of the shape.
ITextPart textPart = paragraph.TextParts[0];
//Modifies the text content of the TextPart.
textPart.Text = "Hello Presentation";
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];
//Modifies the label of the image.
picture.Description = "Replaced alt text";
//Saves the modified PowerPoint Presentation.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);