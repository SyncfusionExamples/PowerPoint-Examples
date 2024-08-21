using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add chart to the slide with position and size.

IPresentationChart chart = slide.Charts.AddChart(100, 10, 700, 500);
//Specify the chart title.
chart.ChartTitle = "Sales Analysis";
//Set chart data - Row1.
chart.ChartData.SetValue(1, 2, "Jan");
chart.ChartData.SetValue(1, 3, "Feb");
chart.ChartData.SetValue(1, 4, "March");

//Set chart data - Row2.
chart.ChartData.SetValue(2, 1, 2010);
chart.ChartData.SetValue(2, 2, 60);
chart.ChartData.SetValue(2, 3, 70);
chart.ChartData.SetValue(2, 4, 80);

//Set chart data - Row3.
chart.ChartData.SetValue(3, 1, 2011);
chart.ChartData.SetValue(3, 2, 80);
chart.ChartData.SetValue(3, 3, 70);
chart.ChartData.SetValue(3, 4, 60);

//Set chart data - Row4.
chart.ChartData.SetValue(4, 1, 2012);
chart.ChartData.SetValue(4, 2, 60);
chart.ChartData.SetValue(4, 3, 70);
chart.ChartData.SetValue(4, 4, 80);

//Create a new chart series with the name.
IOfficeChartSerie seriesJan = chart.Series.Add("Jan");
//Set the data range of chart series – start row, start column, end row, end column.
seriesJan.Values = chart.ChartData[2, 2, 4, 2];
//Create a new chart series with the name.
IOfficeChartSerie seriesFeb = chart.Series.Add("Feb");
//Set the data range of chart series – start row, start column, end row, end column.
seriesFeb.Values = chart.ChartData[2, 3, 4, 3];
//Create a new chart series with the name.
IOfficeChartSerie seriesMarch = chart.Series.Add("March");
//Set the data range of chart series – start row, start column, end row, end column.
seriesMarch.Values = chart.ChartData[2, 4, 4, 4];
//Set the data range of the category axis.
chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 4, 1];
//Specify the chart type.
chart.ChartType = OfficeChartType.Column_Clustered;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
