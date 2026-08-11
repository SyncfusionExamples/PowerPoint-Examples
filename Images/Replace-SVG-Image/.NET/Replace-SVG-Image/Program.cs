using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the icon object from the slide
IPicture icon = slide.Pictures[0];
//Get the new picture as a stream
FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.svg"), FileMode.Open);
//Create an instance for memory stream
MemoryStream memoryStream = new MemoryStream();
//Copy stream to memoryStream
pictureStream.CopyTo(memoryStream);
//Replace the existing icon image with the new image
//SvgData property will return null if it is not an icon
icon.SvgData = memoryStream.ToArray();
//Save the PowerPoint Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Dispose the streams
pictureStream.Dispose();
memoryStream.Dispose();
//Close the Presentation
pptxDoc.Close();