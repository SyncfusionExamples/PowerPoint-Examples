using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Format_Plot_Area
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

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}