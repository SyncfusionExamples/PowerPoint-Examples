using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of a PowerPoint file.
ISlide slide = pptxDoc.Slides[0];
//Get the connector from a slide.
IConnector connector = slide.Shapes[2] as IConnector;
//Remove the connector from slide.
slide.Shapes.Remove(connector);
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));