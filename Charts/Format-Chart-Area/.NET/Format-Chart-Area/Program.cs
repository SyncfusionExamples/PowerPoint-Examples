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
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            //Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
            {
                //Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the chart in slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                //Format the chart area.
                IOfficeChartFrameFormat chartArea = chart.ChartArea;

                //Set border line pattern, color, line weight.
                chartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                chartArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chartArea.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Set fill type and fill colors.
                chartArea.Fill.FillType = OfficeFillType.Gradient;
                chartArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chartArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}