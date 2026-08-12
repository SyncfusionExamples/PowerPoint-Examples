using Syncfusion.Presentation;

//Opens the PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide
ISlide slide = pptxDoc.Slides[0];
//Gets the chart in slide
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Removes the chart from slide
slide.Shapes.Remove(chart as IShape);
//Saves the Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();