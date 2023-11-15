using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Format_Series
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            //Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
            {
                //Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the chart in slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                //Sets chart type and title.
                chart.ChartTitle = "Purchase Details";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.Solid;

                //Sets series type.
                chart.Series[0].SerieType = OfficeChartType.Line_Markers;
                chart.Series[1].SerieType = OfficeChartType.Bar_Clustered;

                chart.PrimaryCategoryAxis.Title = "Products";
                chart.PrimaryValueAxis.Title = "In Dollars";

                // Configure the fill settings for the first series in the chart.
                chart.Series[1].SerieFormat.Fill.FillType = OfficeFillType.Gradient;
                chart.Series[1].SerieFormat.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chart.Series[1].SerieFormat.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                chart.Series[1].SerieFormat.Fill.ForeColor = Syncfusion.Drawing.Color.Red;
                //Customize series border.
                chart.Series[1].SerieFormat.LineProperties.LineColor = Syncfusion.Drawing.Color.Red;
                chart.Series[1].SerieFormat.LineProperties.LinePattern = OfficeChartLinePattern.Dot;
                chart.Series[1].SerieFormat.LineProperties.LineWeight = OfficeChartLineWeight.Hairline;

                //Sets position of legend.
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}