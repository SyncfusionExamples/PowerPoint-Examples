using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide of blank layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to the first slide.
IShape shape = slide.Shapes.AddShape(AutoShapeType.Rectangle, 100, 20, 200, 100);
//Add paragraph into the shape.
IParagraph paragraph = shape.TextBody.AddParagraph();
//Add text to the TextPart.
paragraph.Text = "Syncfusion";
//Set the web hyperlink to the TextPart.
IHyperLink hyperLink = paragraph.TextParts[0].SetHyperlink("http://www.syncfusion.com");
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);