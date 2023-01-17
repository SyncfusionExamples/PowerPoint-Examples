using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
// Add a slide to the PowerPoint file.   
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a rectangle shape on the slide.
IShape rectangle = slide.Shapes.AddShape(AutoShapeType.Rectangle, 420, 250, 100, 100);
//Add connector with specified bounds.
IConnector connector = slide.Shapes.AddConnector(ConnectorType.Straight, 0, 0, 470, 150);
//Connect the beginning point of the connector with rectangle shape.
connector.BeginConnect(rectangle, 0);
//Set the beginning cap of the connector as arrow.
connector.LineFormat.BeginArrowheadStyle = ArrowheadStyle.Arrow;
//Change the connector color.
//Set the connector fill type as solid.
connector.LineFormat.Fill.FillType = FillType.Solid;
//Set the connector solid fill as black.
connector.LineFormat.Fill.SolidFill.Color = ColorObject.Black;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);