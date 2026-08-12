using Syncfusion.Presentation;

//Creates an instance for PowerPoint
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds a normal shape to the slide
slide.Shapes.AddShape(AutoShapeType.Cube, 50, 50, 300, 300);
//Creates an instance for image as stream
FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add picture to the shape collection
IPicture picture = slide.Shapes.AddPicture(imageStream, 373, 83, 526, 382);
//Saves the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
//Closes the stream
imageStream.Close();
//Closes the Presentation
pptxDoc.Close();