using Syncfusion.Presentation;

//Open the specified presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the Ole Object of the slide.
IOleObject oleObject = slide.Shapes[1] as IOleObject;
//Get the data of Ole Image.
byte[] array = oleObject.ImageData;
//Save the extracted Ole data into file system.
using MemoryStream memoryStream = new(array);
using FileStream fileStream = new(Path.GetFullPath(@"Output/OleImage.emf"), FileMode.Create, FileAccess.ReadWrite);
memoryStream.CopyTo(fileStream);