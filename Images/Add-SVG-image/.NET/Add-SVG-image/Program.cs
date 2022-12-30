using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get a SVG image (icon) as stream.
using FileStream svgImageStream = new(@"../../../Data/Image.svg", FileMode.Open);
using FileStream fallbackImageStream = new(@"../../../Data/Image.png", FileMode.Open);
//Add the icon to a slide by specifying its size and position.
IPicture icon = slide.Pictures.AddPicture(svgImageStream, fallbackImageStream, 0, 0, 250, 250);
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);