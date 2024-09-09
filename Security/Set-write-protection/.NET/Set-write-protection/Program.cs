using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add the blank slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add the shape to the slide.
IShape shape = slide.Shapes.AddShape(AutoShapeType.BlockArc, 0, 0, 200, 200);
//Add the paragraph to the shape.
IParagraph paragraph = shape.TextBody.AddParagraph("welcome");
//Set the author name.
pptxDoc.BuiltInDocumentProperties.Author = "Syncfusion";
//Set the write protection for presentation instance.
pptxDoc.SetWriteProtection("MYPASSWORD");
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);