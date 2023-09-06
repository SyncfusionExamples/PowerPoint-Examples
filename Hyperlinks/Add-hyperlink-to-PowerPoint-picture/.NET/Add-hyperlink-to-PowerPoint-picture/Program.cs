using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide of blank layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get a picture as stream.
using FileStream pictureStream = new("../../../Data/Image.jpg", FileMode.Open);
//Add the picture to a slide by specifying its size and position.
IPicture picture = slide.Pictures.AddPicture(pictureStream, 0, 0, 250, 250);
//Set the Email hyperlink to the picture.
IHyperLink hyperLink = (picture as IShape).SetHyperlink("mailto:sales@syncfusion.com");
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);