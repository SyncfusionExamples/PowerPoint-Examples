using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation  
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add chart to the slide with position and size. 
IPresentationChart chart = slide.Charts.AddChart(100, 10, 700, 500);
//Set the chart type as Scatter_Markers. 
chart.ChartType = OfficeChartType.Scatter_Markers;
//Assign data.
chart.DataRange = chart.ChartData[1, 1, 4, 2];
chart.IsSeriesInRows = false;

//Set data to the chart RowIndex, columnIndex, and data. 
chart.ChartData.SetValue(1, 1, "X-Axis");
chart.ChartData.SetValue(1, 2, "Y-Axis");
chart.ChartData.SetValue(2, 1, 1);
chart.ChartData.SetValue(3, 1, 5);
chart.ChartData.SetValue(4, 1, 10);
chart.ChartData.SetValue(2, 2, 10);
chart.ChartData.SetValue(3, 2, 5);
chart.ChartData.SetValue(4, 2, 1);

//Apply chart elements.
//Set chart title.
chart.ChartTitle = "Scatter Markers Chart";
//Set legend.
chart.HasLegend = false;
//Set Datalabels.
IOfficeChartSerie serie = chart.Series[0];
serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
serie.DataPoints.DefaultDataPoint.DataLabels.IsCategoryName = true;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);