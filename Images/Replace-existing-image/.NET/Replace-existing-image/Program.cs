using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];
//Get the new picture as stream.
using FileStream pictureStream = new(@"../../../Data/Image.jpg", FileMode.Open);
//Create instance for memory stream.
using MemoryStream memoryStream = new();
//Copy stream to memoryStream.
pictureStream.CopyTo(memoryStream);
//Replace the existing image with new image.
picture.ImageData = memoryStream.ToArray();
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);