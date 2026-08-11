using Syncfusion.Presentation;

//Open the specified presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide of the Presentation
ISlide slide = pptxDoc.Slides[0];
//Gets the Ole Object of the slide
IOleObject oleObject = slide.Shapes[1] as IOleObject;
//Gets the path of linked Ole Object
string linkOlePath = oleObject.LinkPath;
//Prints the path of linked Ole Object.
System.Console.WriteLine("OleObject link path: "+ linkOlePath);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));