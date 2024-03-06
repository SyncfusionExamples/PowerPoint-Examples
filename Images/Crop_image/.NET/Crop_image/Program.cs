using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation.
using FileStream inputStream = new FileStream(@"../../../Data/Sample.pptx", FileMode.Open, FileAccess.Read);
using IPresentation pptxDoc = Presentation.Open(inputStream);

//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];

//Apply bounding box size and position.
picture.Crop.ContainerWidth = 256.32f;
picture.Crop.ContainerHeight = 153.36f;
picture.Crop.ContainerLeft = 397.44f;
picture.Crop.ContainerTop = 205.2f;

//Apply cropping size and offsets.
picture.Crop.Width = 364.32f;
picture.Crop.Height = 192.24f;
picture.Crop.OffsetX = -27.36f;
picture.Crop.OffsetY = -2.16f;

//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new FileStream("Output.pptx", FileMode.Create);
pptxDoc.Save(outputStream);