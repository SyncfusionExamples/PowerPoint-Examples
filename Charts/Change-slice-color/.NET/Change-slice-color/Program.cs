using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System;
using System.ComponentModel;
using System.Drawing;


namespace Portable
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load or open an PowerPoint Presentation.
            using IPresentation pptxDoc = Presentation.Create();
            //Add a blank slide to the Presentation.
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
            //Add pie chart
            CreatePieChart(slide);
            //Save the output PowerPoint Presentation
            using FileStream outputStream = new(Path.GetFullPath("Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
            pptxDoc.Save(outputStream);
        }
        /// <summary>
        /// Create Pie chart
        /// </summary>
        /// <param name="slide"></param>
        private static void CreatePieChart(ISlide slide)
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

            //Apply chart elements
            //Set Chart Title
            chart.ChartTitle = "Pie Chart";
            //Set Datalabels
            IOfficeChartSerie serie = chart.Series[0];
            serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
            serie.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.BestFit;

            //Change the color of each slice
            serie.DataPoints[0].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Cornsilk;
            serie.DataPoints[1].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Violet;
            serie.DataPoints[2].DataFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Pink;
        }
    }
}