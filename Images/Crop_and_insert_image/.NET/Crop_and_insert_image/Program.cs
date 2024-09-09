using Syncfusion.Presentation;

//Create a new instance of PowerPoint Presentation file.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get a picture as stream.
using FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open);
//Add the picture to a slide by specifying its size and position.
IPicture picture = slide.Pictures.AddPicture(pictureStream, 0, 0, 250, 250);

//Apply bounding box size and position.
picture.Crop.ContainerWidth = 114.48f;
picture.Crop.ContainerHeight = 56.88f;
picture.Crop.ContainerLeft = 94.32f;
picture.Crop.ContainerTop = 128.16f;

//Apply cropping size and offsets.
picture.Crop.Width = 900.72f;
picture.Crop.Height = 74.88f;
picture.Crop.OffsetX = 329.04f;
picture.Crop.OffsetY = -9.36f;

//Save the PowerPoint Presentation.
using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create);
pptxDoc.Save(outputStream);
