
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Add_Drop_Lines
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

                //Set the HasDropLines property to true.
                chart.Series[0].SerieFormat.CommonSerieOptions.HasDropLines = true;

                //Apply formats to DropLines.
                chart.Series[0].SerieFormat.CommonSerieOptions.DropLines.LineColor = Syncfusion.Drawing.Color.Red;
                chart.Series[0].SerieFormat.CommonSerieOptions.DropLines.LinePattern = OfficeChartLinePattern.Dot;
                chart.Series[0].SerieFormat.CommonSerieOptions.DropLines.LineWeight = OfficeChartLineWeight.Hairline;

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