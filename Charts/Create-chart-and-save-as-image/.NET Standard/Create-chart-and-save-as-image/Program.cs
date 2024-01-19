using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;


//Creates a Presentation instance
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a Pie chart
IPresentationChart pieChart = CreatePieChart(slide);
//Initialize the PresentationRenderer
pptxDoc.PresentationRenderer = new PresentationRenderer();
// Converts the chart to image.
Stream image = new FileStream("../../../ChartToImage.jpg", FileMode.Create, FileAccess.ReadWrite);
pptxDoc.PresentationRenderer.ConvertToImage(pieChart, image);
//Closes the presentation
pptxDoc.Close();
image.Close();

IPresentationChart CreatePieChart(ISlide slide)
{
    //Adds chart to the slide with position and size
    IPresentationChart chart = slide.Charts.AddChart(100, 10, 500, 300);
    chart.ChartType = OfficeChartType.Pie;

    //Assign data
    chart.DataRange = chart.ChartData[2, 1, 6, 2];
    chart.IsSeriesInRows = false;

    //Set data to the chart- RowIndex, columnIndex, and data
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

    //Set Chart Title
    chart.ChartTitle = "Pie Chart";

    return chart;
}