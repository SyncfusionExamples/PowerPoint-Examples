using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open the specified presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Gets the first slide of the Presentation
ISlide slide = pptxDoc.Slides[0];
//Gets the Ole Object of the slide
IOleObject oleObject = slide.Shapes[1] as IOleObject;
//Gets the path of linked Ole Object
string linkOlePath = oleObject.LinkPath;
//Prints the path of linked Ole Object.
System.Console.WriteLine("OleObject link path: "+ linkOlePath);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);