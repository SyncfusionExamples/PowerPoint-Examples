using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Chart_Bar_Spacing
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

                //Adding space between bars of different series of single category.
                chart.Series[0].SerieFormat.CommonSerieOptions.Overlap = -40;

                //Adding space between bars of different categories.
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 100;

                //Save the PowerPoint Presentation.
                pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
            }
        }
    }
}