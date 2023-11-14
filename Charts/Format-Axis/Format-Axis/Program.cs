using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Format_Axis
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

                //Set the horizontal (category) axis title.
                chart.PrimaryCategoryAxis.Title = "Months";
                //Set the Vertical (value) axis title.
                chart.PrimaryValueAxis.Title = "Precipitation,in.";
                //Set title for secondary value axis
                chart.SecondaryValueAxis.Title = "Temperature,deg.F";

                //Customize the horizontal category axis.
                chart.PrimaryCategoryAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.PrimaryCategoryAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryCategoryAxis.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Customize the vertical category axis.
                chart.PrimaryValueAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.PrimaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryValueAxis.Border.LineWeight = OfficeChartLineWeight.Narrow;

                //Customize the horizontal category axis font.
                chart.PrimaryCategoryAxis.Font.Color = OfficeKnownColors.Red;
                chart.PrimaryCategoryAxis.Font.FontName = "Calibri";
                chart.PrimaryCategoryAxis.Font.Bold = true;
                chart.PrimaryCategoryAxis.Font.Size = 8;

                //Customize the vertical category axis font.
                chart.PrimaryValueAxis.Font.Color = OfficeKnownColors.Red;
                chart.PrimaryValueAxis.Font.FontName = "Calibri";
                chart.PrimaryValueAxis.Font.Bold = true;
                chart.PrimaryValueAxis.Font.Size = 8;

                //Customize the secondary vertical category axis.
                chart.SecondaryValueAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.SecondaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.SecondaryValueAxis.Border.LineWeight = OfficeChartLineWeight.Narrow;

                //Customize the secondary vertical category axis font.
                chart.SecondaryValueAxis.Font.Color = OfficeKnownColors.Red;
                chart.SecondaryValueAxis.Font.FontName = "Calibri";
                chart.SecondaryValueAxis.Font.Bold = true;
                chart.SecondaryValueAxis.Font.Size = 8;

                //Axis title area text angle rotation.
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 270;

                //Maximum value in the axis.
                chart.PrimaryValueAxis.MaximumValue = 15;
                chart.PrimaryValueAxis.MinimumValue = 0;
                //Number format for axis.
                chart.PrimaryValueAxis.NumberFormat = "0.0";

                //Hiding major gridlines.
                chart.PrimaryValueAxis.HasMajorGridLines = true;

                //Showing minor gridlines.
                chart.PrimaryValueAxis.HasMinorGridLines = false;

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