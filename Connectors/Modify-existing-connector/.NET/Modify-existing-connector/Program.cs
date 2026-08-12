using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of a PowerPoint file.
ISlide slide = pptxDoc.Slides[0];
//Get the connector from a slide.
IConnector connector = slide.Shapes[2] as IConnector;
//Set the begin cap for the connector.
connector.LineFormat.BeginArrowheadStyle = ArrowheadStyle.ArrowOpen;
//Set the line format for the connector.
connector.LineFormat.DashStyle = LineDashStyle.DashDotDot;
//Disconnect the end connection of the connector if end point get connected.
if (connector.EndConnectedShape != null)
    connector.EndDisconnect();
//Insert a triangle shape into slide.
IShape triangle = slide.Shapes.AddShape(AutoShapeType.IsoscelesTriangle, 600, 200, 150, 150);
//Declare the end connection site index.
int connectionSiteIndex = 4;
//Reconnect the end point of connector with triangle shape if its connection site count is greater than 4.
if (connectionSiteIndex < triangle.ConnectionSiteCount)
    connector.EndConnect(triangle, connectionSiteIndex);
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));