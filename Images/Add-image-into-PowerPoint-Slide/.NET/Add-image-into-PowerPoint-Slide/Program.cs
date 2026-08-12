using Syncfusion.Presentation;

//Create an instance of Presentation
IPresentation pptxDoc = Presentation.Create();
//Add a blank slide.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get a picture as a stream.
FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add the picture to a slide by specifying its size and position.
IPicture picture = slide.Pictures.AddPicture(pictureStream, 0, 0, 250, 250);
//Save the PowerPoint Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Dispose the image stream
pictureStream.Dispose();
//Close the Presentation
pptxDoc.Close();