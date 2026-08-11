using Syncfusion.OfficeChart;
using Syncfusion.Presentation;


namespace Format_Chart_Title
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                // Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                // Gets the chart in the slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                // Set the chart title.
                chart.ChartTitle = "Purchase Details";

                // Customize the chart title area.
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Color = OfficeKnownColors.Black;
                chart.ChartTitleArea.Underline = OfficeUnderline.WavyHeavy;
                chart.ChartTitleArea.Size = 14;

                // Position the chart title area using Layout (values in points).
                chart.ChartTitleArea.Layout.Top = 10;
                chart.ChartTitleArea.Layout.Left = 10;

                // Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}