using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
IPresentationChart chart = slide1.Charts.AddChart(50, 50, 600, 400);
chart.ChartTitle = "Test Scores";
chart.ChartType = OfficeChartType.BoxAndWhisker;
//Assign data.
chart.DataRange = chart.ChartData[1, 1, 16, 4];
chart.IsSeriesInRows = false;
//Set data to the chart RowIndex, columnIndex, and data.
SetChartData(chart);

//Box and Whisker settings on first series.
IOfficeChartSerie seriesA = chart.Series[0];
seriesA.SerieFormat.ShowInnerPoints = false;
seriesA.SerieFormat.ShowOutlierPoints = true;
seriesA.SerieFormat.ShowMeanMarkers = true;
seriesA.SerieFormat.ShowMeanLine = false;
seriesA.SerieFormat.QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

//Box and Whisker settings on second series.
IOfficeChartSerie seriesB = chart.Series[1];
seriesB.SerieFormat.ShowInnerPoints = false;
seriesB.SerieFormat.ShowOutlierPoints = true;
seriesB.SerieFormat.ShowMeanMarkers = true;
seriesB.SerieFormat.ShowMeanLine = false;
seriesB.SerieFormat.QuartileCalculationType = QuartileCalculation.InclusiveMedian;

//Box and Whisker settings on third series.
IOfficeChartSerie seriesC = chart.Series[2];
seriesC.SerieFormat.ShowInnerPoints = false;
seriesC.SerieFormat.ShowOutlierPoints = true;
seriesC.SerieFormat.ShowMeanMarkers = true;
seriesC.SerieFormat.ShowMeanLine = false;
seriesC.SerieFormat.QuartileCalculationType = QuartileCalculation.ExclusiveMedian;

using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream); 

/// <summary>
/// Set the values for the chart
/// </summary>
/// <param name="chart">Represent the instance of the Presentation chart</param>
static void SetChartData(IPresentationChart chart)
{
    chart.ChartData.SetValue(1, 1, "Course");
    chart.ChartData.SetValue(1, 2, "SchoolA");
    chart.ChartData.SetValue(1, 3, "SchoolB");
    chart.ChartData.SetValue(1, 4, "SchoolC");

    chart.ChartData.SetValue(2, 1, "English");
    chart.ChartData.SetValue(2, 2, 63);
    chart.ChartData.SetValue(2, 3, 53);
    chart.ChartData.SetValue(2, 4, 45);

    chart.ChartData.SetValue(3, 1, "Physics");
    chart.ChartData.SetValue(3, 2, 61);
    chart.ChartData.SetValue(3, 3, 55);
    chart.ChartData.SetValue(3, 4, 65);

    chart.ChartData.SetValue(4, 1, "English");
    chart.ChartData.SetValue(4, 2, 63);
    chart.ChartData.SetValue(4, 3, 50);
    chart.ChartData.SetValue(4, 4, 65);

    chart.ChartData.SetValue(5, 1, "Math");
    chart.ChartData.SetValue(5, 2, 62);
    chart.ChartData.SetValue(5, 3, 51);
    chart.ChartData.SetValue(5, 4, 64);

    chart.ChartData.SetValue(6, 1, "English");
    chart.ChartData.SetValue(6, 2, 46);
    chart.ChartData.SetValue(6, 3, 53);
    chart.ChartData.SetValue(6, 4, 66);

    chart.ChartData.SetValue(7, 1, "English");
    chart.ChartData.SetValue(7, 2, 58);
    chart.ChartData.SetValue(7, 3, 56);
    chart.ChartData.SetValue(7, 4, 67);

    chart.ChartData.SetValue(8, 1, "Math");
    chart.ChartData.SetValue(8, 2, 62);
    chart.ChartData.SetValue(8, 3, 53);
    chart.ChartData.SetValue(8, 4, 66);

    chart.ChartData.SetValue(9, 1, "Math");
    chart.ChartData.SetValue(9, 2, 62);
    chart.ChartData.SetValue(9, 3, 53);
    chart.ChartData.SetValue(9, 4, 66);

    chart.ChartData.SetValue(10, 1, "English");
    chart.ChartData.SetValue(10, 2, 63);
    chart.ChartData.SetValue(10, 3, 54);
    chart.ChartData.SetValue(10, 4, 64);

    chart.ChartData.SetValue(11, 1, "English");
    chart.ChartData.SetValue(11, 2, 63);
    chart.ChartData.SetValue(11, 3, 52);
    chart.ChartData.SetValue(11, 4, 67);

    chart.ChartData.SetValue(12, 1, "Physics");
    chart.ChartData.SetValue(12, 2, 60);
    chart.ChartData.SetValue(12, 3, 56);
    chart.ChartData.SetValue(12, 4, 64);

    chart.ChartData.SetValue(13, 1, "English");
    chart.ChartData.SetValue(13, 2, 60);
    chart.ChartData.SetValue(13, 3, 56);
    chart.ChartData.SetValue(13, 4, 64);

    chart.ChartData.SetValue(14, 1, "Math");
    chart.ChartData.SetValue(14, 2, 61);
    chart.ChartData.SetValue(14, 3, 56);
    chart.ChartData.SetValue(14, 4, 45);

    chart.ChartData.SetValue(15, 1, "Math");
    chart.ChartData.SetValue(15, 2, 63);
    chart.ChartData.SetValue(15, 3, 58);
    chart.ChartData.SetValue(15, 4, 64);

    chart.ChartData.SetValue(16, 1, "English");
    chart.ChartData.SetValue(16, 2, 59);
    chart.ChartData.SetValue(16, 3, 54);
    chart.ChartData.SetValue(16, 4, 65);
}