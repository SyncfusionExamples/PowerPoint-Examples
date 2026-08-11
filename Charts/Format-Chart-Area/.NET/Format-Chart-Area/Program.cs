using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.Reflection.Metadata;
using System;


namespace Format_Chart_Area
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

                //Format the chart area.
                IOfficeChartFrameFormat chartArea = chart.ChartArea;

                //Set border line pattern, color, and line weight.
                chartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                chartArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chartArea.Border.LineWeight = OfficeChartLineWeight.Hairline;
                //Set fill type and fill colors.
                chartArea.Fill.FillType = OfficeFillType.Gradient;
                chartArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chartArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;

                //Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}