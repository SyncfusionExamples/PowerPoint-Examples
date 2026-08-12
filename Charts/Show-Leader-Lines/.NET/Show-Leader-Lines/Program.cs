
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
                //Gets the first slide.
                ISlide slide = pptxDoc.Slides[0];
                //Gets the chart in slide.
                IPresentationChart chart = slide.Charts[0];

                //Show leader lines enabled.
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;

                //Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}