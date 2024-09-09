using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide.
slide.Shapes.AddShape(AutoShapeType.Cube, 50, 200, 300, 300);
//Get a picture as stream.
using FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add picture to the shape collection.
IPicture picture = slide.Shapes.AddPicture(imageStream, 373, 83, 526, 382);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);