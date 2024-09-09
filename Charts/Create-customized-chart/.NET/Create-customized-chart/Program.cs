using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a new chart in the slide by specifying its position and size as parameters.
IPresentationChart chart = slide.Charts.AddChart(100, 80, 500, 350);
chart.ChartTitle = "Sales comparison";
chart.ChartTitleArea.Bold = true;

//Set the data for chart– RowIndex, columnIndex and data.
chart.ChartData.SetValue(1, 1, "Month");
chart.ChartData.SetValue(2, 1, "July");
chart.ChartData.SetValue(3, 1, "August");
chart.ChartData.SetValue(4, 1, "September");
chart.ChartData.SetValue(5, 1, "October");
chart.ChartData.SetValue(6, 1, "November");
chart.ChartData.SetValue(7, 1, "December");

chart.ChartData.SetValue(1, 2, "2013");
chart.ChartData.SetValue(2, 2, 35);
chart.ChartData.SetValue(3, 2, 47);
chart.ChartData.SetValue(4, 2, 30);
chart.ChartData.SetValue(5, 2, 29);
chart.ChartData.SetValue(6, 2, 25);
chart.ChartData.SetValue(7, 2, 30);

chart.ChartData.SetValue(1, 3, "2014");
chart.ChartData.SetValue(2, 3, 30);
chart.ChartData.SetValue(3, 3, 25);
chart.ChartData.SetValue(4, 3, 29);
chart.ChartData.SetValue(5, 3, 35);
chart.ChartData.SetValue(6, 3, 38);
chart.ChartData.SetValue(7, 3, 32);

chart.ChartData.SetValue(1, 4, "2015");
chart.ChartData.SetValue(2, 4, 35);
chart.ChartData.SetValue(3, 4, 37);
chart.ChartData.SetValue(4, 4, 30);
chart.ChartData.SetValue(5, 4, 29);
chart.ChartData.SetValue(6, 4, 25);
chart.ChartData.SetValue(7, 4, 30);

//Create a new ChartSerie with the name.
IOfficeChartSerie serie2013 = chart.Series.Add("2013");
//Set the data range of chart series start row, start column, end row, end column.
serie2013.Values = chart.ChartData[2, 2, 7, 2];
serie2013.SerieType = OfficeChartType.Bar_Clustered;
IOfficeChartSerie serie2014 = chart.Series.Add("2014");
serie2014.Values = chart.ChartData[2, 3, 7, 3];
serie2014.SerieType = OfficeChartType.Scatter_Line_Markers;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
