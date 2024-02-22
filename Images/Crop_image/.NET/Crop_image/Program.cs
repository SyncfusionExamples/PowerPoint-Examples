using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation.
using FileStream inputStream = new FileStream(@"../../../Data/Sample.pptx", FileMode.Open, FileAccess.Read);
using IPresentation pptxDoc = Presentation.Open(inputStream);

//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];

//Apply cropping to the picture.
picture.Crop.ContainerWidth = 345;
picture.Crop.ContainerHeight = 220;
picture.Crop.ContainerLeft = 324;
picture.Crop.ContainerTop = 220;

picture.Crop.Width = 288;
picture.Crop.Height = 144;
picture.Crop.OffsetX = -35;
picture.Crop.OffsetY = -65;

//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new FileStream("Output.pptx", FileMode.Create);
pptxDoc.Save(outputStream);