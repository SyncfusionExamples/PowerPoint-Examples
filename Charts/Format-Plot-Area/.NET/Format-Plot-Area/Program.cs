using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Format_Plot_Area
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                //Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the chart in the slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                //Plot Area.
                IOfficeChartFrameFormat chartPlotArea = chart.PlotArea;

                //Plot area border settings - line pattern, color, weight.
                chartPlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                chartPlotArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chartPlotArea.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Set fill type and color.
                chartPlotArea.Fill.FillType = OfficeFillType.Gradient;
                chartPlotArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chartPlotArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;

                //Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}