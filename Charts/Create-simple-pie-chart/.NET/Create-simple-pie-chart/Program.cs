using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Create an instance of presentation.
using IPresentation presentation = Presentation.Create();
//Add a blank slide to the presentation.
ISlide slide = presentation.Slides.Add(SlideLayoutType.Blank);
//Add a chart to the slide with position and size.
IPresentationChart chart = slide.Charts.AddChart(100, 10, 700, 500);
//Select the chart type.
chart.ChartType = OfficeChartType.Pie;
//Assign data range.
chart.DataRange = chart.ChartData[1, 1, 6, 2];
chart.IsSeriesInRows = false;
//Set the values for the chart data.
chart.ChartData.SetValue(1, 1, "Food");
chart.ChartData.SetValue(2, 1, "Fruits");
chart.ChartData.SetValue(3, 1, "Vegetables");
chart.ChartData.SetValue(4, 1, "Dairy");
chart.ChartData.SetValue(5, 1, "Protein");
chart.ChartData.SetValue(6, 1, "Grains");
chart.ChartData.SetValue(1, 2, "Percentage");
chart.ChartData.SetValue(2, 2, 36);
chart.ChartData.SetValue(3, 2, 14);
chart.ChartData.SetValue(4, 2, 13);
chart.ChartData.SetValue(5, 2, 28);
chart.ChartData.SetValue(6, 2, 9);
//Set data labels.
chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
//Save the presentation.
using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
presentation.Save(outputStream);
