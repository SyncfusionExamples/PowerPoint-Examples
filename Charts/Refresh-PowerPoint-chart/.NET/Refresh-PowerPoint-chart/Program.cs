using Syncfusion.Presentation;

//Opens the PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide
ISlide slide = pptxDoc.Slides[0];
//Gets the chart in slide
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Refreshes the chart data. Set true to evaluate Excel formulas before refreshing,
//or false to refresh only the data without evaluating formulas.
chart.Refresh(false);
//Saves the Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();