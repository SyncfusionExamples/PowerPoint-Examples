using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the icon object from the slide
IPicture icon = slide.Pictures[0];
//Get the new picture as stream.
using FileStream pictureStream = new(Path.GetFullPath(@"Data/Image.svg"), FileMode.Open);
//Create instance for memory stream
using MemoryStream memoryStream = new();
//Copy stream to memoryStream.
pictureStream.CopyTo(memoryStream);
//Replace the existing icon image with new image
//SvgData property will return null, if it is not an icon
icon.SvgData = memoryStream.ToArray();
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);