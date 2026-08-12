using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of a PowerPoint file.
ISlide slide = pptxDoc.Slides[0];
//Get the rectangle shape from a slide.
IShape rectangle = slide.Shapes[0] as IShape;
//Get the connector from a slide.
IConnector connector = slide.Shapes[2] as IConnector;
//Change the X and Y position of the rectangle.
rectangle.Left = 600;
rectangle.Top = 200;
//Update the connector to connect with previously updated shape.
connector.Update();
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));