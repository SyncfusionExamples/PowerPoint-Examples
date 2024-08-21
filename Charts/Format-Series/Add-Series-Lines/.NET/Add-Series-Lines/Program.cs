      
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Add_Series_Lines
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            //Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
            {
                //Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the chart in slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                //Set Chart Type
                chart.ChartType = OfficeChartType.Bar_Stacked;
                //Set HasSeriesLines  property to true.
                chart.Series[0].SerieFormat.CommonSerieOptions.HasSeriesLines = true;

                //Apply formats to Series Lines.
                chart.Series[0].SerieFormat.CommonSerieOptions.PieSeriesLine.LineColor = Syncfusion.Drawing.Color.Red;
                chart.Series[0].SerieFormat.CommonSerieOptions.PieSeriesLine.LinePattern = OfficeChartLinePattern.Solid;
                chart.Series[0].SerieFormat.CommonSerieOptions.PieSeriesLine.LineWeight = OfficeChartLineWeight.Medium;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}