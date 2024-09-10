using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a group shape to the slide.
IGroupShape groupShape = slide.GroupShapes.AddGroupShape(20, 20, 450, 300);
//Add a TextBox to the group shape.
groupShape.Shapes.AddTextBox(30, 25, 100, 100).TextBody.AddParagraph("My TextBox");
//Get a picture as stream.
using FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add a picture to the group shape.
groupShape.Shapes.AddPicture(pictureStream, 40, 100, 100, 100);
//Add a shape to the group shape.
groupShape.Shapes.AddShape(AutoShapeType.Rectangle, 200, 200, 90, 30);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);