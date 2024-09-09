using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide.
ISlide slide = pptxDoc.Slides[0];
//Get the chart in slide.
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Remove the chart from slide.
slide.Shapes.Remove(chart as IShape); 
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);