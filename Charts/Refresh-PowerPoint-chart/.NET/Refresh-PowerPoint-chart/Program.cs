using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide.
ISlide slide = pptxDoc.Slides[0];
//Get the chart in slide.
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
// Refreshes the chart data. Set `true` to evaluate Excel formulas before refreshing,
// or `false` to refresh only the data without evaluating formulas.
chart.Refresh(false);
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);