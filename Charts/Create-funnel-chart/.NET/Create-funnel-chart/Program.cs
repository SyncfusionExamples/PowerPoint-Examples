using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide to Presentation.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Create a chart.
IPresentationChart chart = slide1.Charts.AddChart(30, 50, 600, 300);
//Set chart type as Funnel.
chart.ChartType = OfficeChartType.Funnel;
//Set the chart title.
chart.ChartTitle = "Funnel";
//Assign data.
chart.DataRange = chart.ChartData[1, 1, 6, 2];
chart.IsSeriesInRows = false;
//Set data.
chart.ChartData.SetValue(1, 1, "Web sales");
chart.ChartData.SetValue(1, 2, "Users count");
chart.ChartData.SetValue(2, 1, "Website Visits");
chart.ChartData.SetValue(2, 2, "15600");
chart.ChartData.SetValue(3, 1, "Downloads");
chart.ChartData.SetValue(3, 2, "8000");
chart.ChartData.SetValue(4, 1, "Requested price list");
chart.ChartData.SetValue(4, 2, "6000");
chart.ChartData.SetValue(5, 1, "Invoice sent");
chart.ChartData.SetValue(5, 2, "2000");
chart.ChartData.SetValue(6, 1, "Finalized");
chart.ChartData.SetValue(6, 2, "1000");
//Formatting the legend and data label option.
chart.HasLegend = false;
IOfficeChartSerie serie = chart.Series[0];
chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
