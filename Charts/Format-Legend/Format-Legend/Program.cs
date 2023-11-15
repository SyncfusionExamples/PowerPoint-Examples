
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Format_Legend
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

                //Enable the legend.
                chart.HasLegend = true;

                //Set the position of legend.
                chart.Legend.Position = OfficeLegendPosition.Right;

                //Legend without overlapping the chart.
                chart.Legend.IncludeInLayout = true;
                chart.Legend.FrameFormat.Border.AutoFormat = false;
                chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                chart.Legend.FrameFormat.Border.LineColor = Syncfusion.Drawing.Color.Black;
                chart.Legend.FrameFormat.Border.LinePattern = OfficeChartLinePattern.DashDot;
                chart.Legend.FrameFormat.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Set the legend's text area formatting - font name, weight, color, size.
                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.Color = OfficeKnownColors.Pink;
                chart.Legend.TextArea.FontName = "Times New Roman";
                chart.Legend.TextArea.Size = 10;
                chart.Legend.TextArea.Strikethrough = false;

                //View legend in vertical.
                chart.Legend.IsVerticalLegend = true;

                //Modifies the legend entry.
                chart.Legend.LegendEntries[0].IsDeleted = true;

                //Manually resizing chart legend area using Layout.
                chart.Legend.Layout.Left = 0.2;
                chart.Legend.Layout.Top = 5;
                chart.Legend.Layout.Width = 40;
                chart.Legend.Layout.Height = 40;

                //Legend without overlapping the chart.
                chart.Legend.IncludeInLayout = true;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}