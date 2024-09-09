using Syncfusion.OfficeChart;
using Syncfusion.Presentation;
using System.ComponentModel;


namespace Chart_Bar_Spacing
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

                //Adding space between bars of different series of single category.
                chart.Series[0].SerieFormat.CommonSerieOptions.Overlap = -40;

                //Adding space between bars of different categories.
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 100;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the PowerPoint Presentation.
                    pptxDoc.Save(outputStream);
                }
            }
        }
    }
}