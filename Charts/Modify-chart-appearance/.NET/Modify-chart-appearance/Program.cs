using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide.
ISlide slide = pptxDoc.Slides[0];
//Get the chart in slide.
IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
//Modify the chart height.
chart.Height = 500;
//Modifies the chart width.
chart.Width = 700;
//Change the title.
chart.ChartTitle = "New title";
//Changes the series name of first chart series.
chart.Series[0].Name = "Modified series name";
//Hide the category labels.
chart.CategoryLabelLevel = OfficeCategoriesLabelLevel.CategoriesLabelLevelNone;
//Show Data Table.
chart.HasDataTable = true;

//Format Chart Area.
IOfficeChartFrameFormat chartArea = chart.ChartArea;
//Chart Area Border Settings.
//Style
chartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
//Color
chartArea.Border.LineColor = Color.Blue;
//Weight
chartArea.Border.LineWeight = OfficeChartLineWeight.Hairline;
//Chart Area Settings.
//Fill Effects
chartArea.Fill.FillType = OfficeFillType.Gradient;
//Two Color
chartArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
//Sets two colors
chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
chartArea.Fill.ForeColor = Color.White;

//Plot Area
IOfficeChartFrameFormat chartPlotArea = chart.PlotArea;
//Plots Area Border Settings
//Style
chartPlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
//Color
chartPlotArea.Border.LineColor = Color.Blue;
//Weight
chartPlotArea.Border.LineWeight = OfficeChartLineWeight.Hairline;
//Fill Effects
chartPlotArea.Fill.FillType = OfficeFillType.Gradient;
//Two Color
chartPlotArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
//Set two colors.
chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
chartPlotArea.Fill.ForeColor = Color.White;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);