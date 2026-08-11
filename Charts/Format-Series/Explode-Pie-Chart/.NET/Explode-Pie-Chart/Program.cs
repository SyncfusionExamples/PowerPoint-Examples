using Syncfusion.OfficeChart;
using Syncfusion.Presentation;

namespace Explode_Pie_Chart
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
                IPresentationChart chart = slide.Shapes[0] as IPresentationChart;

                //Exploding the pie chart to 40%.
                chart.Series[0].SerieFormat.Percent = 40;

                //Sets position of legend.
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                //Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}