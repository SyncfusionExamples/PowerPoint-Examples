using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide to the PowerPoint file.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a rectangle shape on the slide.
IShape rectangle = slide.Shapes.AddShape(AutoShapeType.Rectangle, 200, 300, 100, 100);
//Add an oval shape on the slide.
IShape oval = slide.Shapes.AddShape(AutoShapeType.Oval, 400, 10, 100, 100);
//Add elbow connector on the slide and connect the end points of connector with specified port positions 0 and 4 of the beginning and end shapes.
IConnector connector = slide.Shapes.AddConnector(ConnectorType.Elbow, rectangle, 0, oval, 4);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);