using Syncfusion.OfficeChart;
using Syncfusion.Presentation;


namespace Format_Chart_Title
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

                // Set the chart title.
                chart.ChartTitle = "Purchase Details";

                // Customize chart title area.
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Color = OfficeKnownColors.Black;
                chart.ChartTitleArea.Underline = OfficeUnderline.WavyHeavy;

                //Manually resizing chart title area using Layout.
                chart.ChartTitleArea.Layout.Left = 5;

                //Enable legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Right;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
                // Open the PowerPoint Presentation located at the specified path using the default associated program.
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Result.pptx"))
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}