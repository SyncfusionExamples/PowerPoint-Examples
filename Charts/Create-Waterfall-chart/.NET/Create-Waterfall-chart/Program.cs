using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the presentation.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Create a chart.
IPresentationChart chart = slide1.Charts.AddChart(50, 50, 700, 400);
//Set chart type as Waterfall.
chart.ChartType = OfficeChartType.WaterFall;

//Set data range .
chart.DataRange = chart.ChartData[1, 1, 8, 2];
chart.IsSeriesInRows = false;
chart.ChartData.SetValue(2, 1, "Start");
chart.ChartData.SetValue(2, 2, 120000);
chart.ChartData.SetValue(3, 1, "Product Revenue");
chart.ChartData.SetValue(3, 2, 570000);
chart.ChartData.SetValue(4, 1, "Service Revenue");
chart.ChartData.SetValue(4, 2, 230000);
chart.ChartData.SetValue(5, 1, "Positive Balance");
chart.ChartData.SetValue(5, 2, 920000);
chart.ChartData.SetValue(6, 1, "Fixed Costs");
chart.ChartData.SetValue(6, 2, -345000);
chart.ChartData.SetValue(7, 1, "Variable Costs");
chart.ChartData.SetValue(7, 2, -230000);
chart.ChartData.SetValue(8, 1, "Total");
chart.ChartData.SetValue(8, 2, 345000);

//Data point settings as total in chart.
IOfficeChartSerie series = chart.Series[0];
chart.Series[0].DataPoints[3].SetAsTotal = true;
chart.Series[0].DataPoints[6].SetAsTotal = true;
//Showing the connector lines between data points.
chart.Series[0].SerieFormat.ShowConnectorLines = true;
//Set the chart title.
chart.ChartTitle = "Company Profit (in USD)";
//Formatting data label and legend option.
chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
chart.Legend.Position = OfficeLegendPosition.Right;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
