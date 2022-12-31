using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide to Presentation.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);

IPresentationChart chart = slide1.Charts.AddChart(50, 50, 500, 400);
//Set chart type as Pareto.
chart.ChartType = OfficeChartType.Pareto;
//Set data range.
chart.DataRange = chart.ChartData[2, 1, 8, 2];
chart.ChartData.SetValue(2, 1, "Rent");
chart.ChartData.SetValue(2, 2, 2300);
chart.ChartData.SetValue(3, 1, "Car payment");
chart.ChartData.SetValue(3, 2, 1200);
chart.ChartData.SetValue(4, 1, "Groceries");
chart.ChartData.SetValue(4, 2, 900);
chart.ChartData.SetValue(5, 1, "Electricity");
chart.ChartData.SetValue(5, 2, 600);
chart.ChartData.SetValue(6, 1, "Gas");
chart.ChartData.SetValue(6, 2, 500);
chart.ChartData.SetValue(7, 1, "Cable");
chart.ChartData.SetValue(7, 2, 300);
chart.ChartData.SetValue(8, 1, "Mobile");
chart.ChartData.SetValue(8, 2, 200);

//Set category values as bin values  .
chart.PrimaryCategoryAxis.IsBinningByCategory = true;
//Formatting Pareto line.     
chart.Series[0].ParetoLineFormat.LineProperties.ColorIndex = OfficeKnownColors.Bright_green;
//Gap width settings.
chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 6;
//Set the chart title.
chart.ChartTitle = "Expenses";
//Hiding the legend.
chart.HasLegend = false;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
