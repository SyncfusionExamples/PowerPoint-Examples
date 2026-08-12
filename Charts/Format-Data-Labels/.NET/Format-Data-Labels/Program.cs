
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Format_Data_Labels
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                // Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                // Gets the chart in the slide.
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
                for (int i = 0; i < chart.Series.Count; i++)
                {
                    // Enable the data labels in the chart.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                    // Set the font size of the data labels.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                    // Change the color of the data labels.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Color = OfficeKnownColors.Black;
                    // Make the data labels bold.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Bold = true;
                    // Set the position of data labels for the first series.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                }
                // Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}