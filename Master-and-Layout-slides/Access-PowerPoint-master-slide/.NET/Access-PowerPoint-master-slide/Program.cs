using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Access the first master slide in PowerPoint file.
IMasterSlide masterSlide = pptxDoc.Masters[0];
//Get the first shape name from the master slide.
string shapeName = masterSlide.Shapes[0].ShapeName;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);