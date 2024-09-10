using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open the specified presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide of the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the Ole Object of the slide.
IOleObject oleObject = slide.Shapes[2] as IOleObject;
//Get the file data of embedded Ole Object.
byte[] array = oleObject.ObjectData;
//Get the file Name of OLE Object.
string outputFile = oleObject.FileName;
//Save the extracted Ole data into file system.
using MemoryStream memoryStream = new(array);
using FileStream fileStream = new(Path.GetFullPath(@"Output/"+outputFile), FileMode.Create, FileAccess.ReadWrite);
memoryStream.CopyTo(fileStream);