using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);

IPresentationChart chart = slide1.Charts.AddChart(50, 50, 500, 400);
chart.ChartType = OfficeChartType.Histogram;
chart.DataRange = chart.ChartData[2, 1, 15, 1];
chart.ChartData.SetValue(1, 1, "Student Heights");
chart.ChartData.SetValue(2, 1, 130);
chart.ChartData.SetValue(3, 1, 132);
chart.ChartData.SetValue(4, 1, 159);
chart.ChartData.SetValue(5, 1, 163);
chart.ChartData.SetValue(6, 1, 140);
chart.ChartData.SetValue(7, 1, 155);
chart.ChartData.SetValue(8, 1, 139);
chart.ChartData.SetValue(9, 1, 143);
chart.ChartData.SetValue(10, 1, 153);
chart.ChartData.SetValue(11, 1, 165);
chart.ChartData.SetValue(12, 1, 153);
chart.ChartData.SetValue(13, 1, 149);
chart.ChartData.SetValue(14, 1, 154);
chart.ChartData.SetValue(15, 1, 162);

//Category axis bin settings.
chart.PrimaryCategoryAxis.BinWidth = 8;
//Gap width settings.
chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 6;
//Set the chart title and axis title.
chart.ChartTitle = "Height Data";
chart.PrimaryValueAxis.Title = "Number of students";
chart.PrimaryCategoryAxis.Title = "Height";
//Hiding the legend.
chart.HasLegend = false;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);