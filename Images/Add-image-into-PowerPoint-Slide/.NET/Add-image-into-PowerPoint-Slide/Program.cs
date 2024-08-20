using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get a picture as stream.
using FileStream pictureStream = new(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add the picture to a slide by specifying its size and position.
IPicture picture = slide.Pictures.AddPicture(pictureStream, 0, 0, 250, 250);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);