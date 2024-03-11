using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation.
using FileStream inputStream = new FileStream(@"../../../Data/Sample.pptx", FileMode.Open, FileAccess.Read);
using IPresentation pptxDoc = Presentation.Open(inputStream);

//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];

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

//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new FileStream("Output.pptx", FileMode.Create);
pptxDoc.Save(outputStream);