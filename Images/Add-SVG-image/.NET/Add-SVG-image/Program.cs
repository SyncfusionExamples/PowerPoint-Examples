using Syncfusion.Presentation;

//Create an instance of Presentation
IPresentation pptxDoc = Presentation.Create();
//Add a blank slide
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get an SVG image (icon) as a stream
FileStream svgImageStream = new FileStream(Path.GetFullPath(@"Data/Image.svg"), FileMode.Open);
//Get a fallback image as a stream
FileStream fallbackImageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open);
//Add the icon to a slide by specifying its size and position
IPicture icon = slide.Pictures.AddPicture(svgImageStream, fallbackImageStream, 0, 0, 250, 250);
//Save the PowerPoint Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Dispose the fallback image stream
fallbackImageStream.Dispose();
//Dispose the SVG image stream
svgImageStream.Dispose();
//Close the Presentation
pptxDoc.Close();