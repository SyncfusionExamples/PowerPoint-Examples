
using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Format_Data_Labels
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

                for (int i = 0; i < chart.Series.Count; i++)
                {
                    //Enable the datalabel in chart.
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

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}