using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation presentation = Presentation.Create();
//Add slide to Presentation.
ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
//Add textbox to slide.
IShape shape = slide.Shapes.AddTextBox(100, 30, 200, 300);
//Add a paragraph with text content.
IParagraph paragraph = shape.TextBody.AddParagraph("Password Protected.");
//Protects the file with password.
presentation.Encrypt("PASSWORD!@1#$");
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
presentation.Save(outputStream);