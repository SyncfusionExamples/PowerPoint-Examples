using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];
//Get the new picture as a stream.
FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Create an instance for memory stream
MemoryStream memoryStream = new MemoryStream();
//Copy stream to memoryStream.
pictureStream.CopyTo(memoryStream);
//Replace the existing image with the new image.
picture.ImageData = memoryStream.ToArray();
//Save the PowerPoint Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Dispose the streams
pictureStream.Dispose();
memoryStream.Dispose();
//Close the Presentation
pptxDoc.Close();