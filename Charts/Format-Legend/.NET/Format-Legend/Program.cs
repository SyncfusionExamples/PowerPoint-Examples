
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Format_Legend
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                // Get the first slide.
                ISlide slide = pptxDoc.Slides[0];
                // Get the chart in the slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                // Enable the legend.
                chart.HasLegend = true;

                // Set the position of the legend.
                chart.Legend.Position = OfficeLegendPosition.Right;

                // Include the legend in the chart layout and apply the legend border.
                chart.Legend.IncludeInLayout = true;
                chart.Legend.FrameFormat.Border.AutoFormat = false;
                chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                chart.Legend.FrameFormat.Border.LineColor = Syncfusion.Drawing.Color.Black;
                chart.Legend.FrameFormat.Border.LinePattern = OfficeChartLinePattern.DashDot;
                chart.Legend.FrameFormat.Border.LineWeight = OfficeChartLineWeight.Hairline;

                // Set the legend's text area formatting - font name, weight, color, size.
                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.Color = OfficeKnownColors.Pink;
                chart.Legend.TextArea.FontName = "Times New Roman";
                chart.Legend.TextArea.Size = 10;
                chart.Legend.TextArea.Strikethrough = false;

                // Display the legend vertically.
                chart.Legend.IsVerticalLegend = true;

                // Delete the legend entry.
                chart.Legend.LegendEntries[0].IsDeleted = true;

                // Resize the legend area (values in points).
                chart.Legend.Layout.Left = 0.2;
                chart.Legend.Layout.Top = 5;
                chart.Legend.Layout.Width = 40;
                chart.Legend.Layout.Height = 40;

                // Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}