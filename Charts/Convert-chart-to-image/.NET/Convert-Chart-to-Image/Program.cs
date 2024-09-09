using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;

namespace Convert_Chart_to_Image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Loads or open an PowerPoint Presentation
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.pptx"), FileMode.Open);
            IPresentation pptxDoc = Presentation.Open(inputStream);
            //Initialize the PresentationRenderer
            pptxDoc.PresentationRenderer = new PresentationRenderer();
            //Gets the first instance of chart from slide
            IPresentationChart chart = pptxDoc.Slides[0].Charts[0];
            // Converts the chart to image.
            Stream image = new FileStream(Path.GetFullPath(@"Output/Output.jpg"), FileMode.Create, FileAccess.ReadWrite);
            pptxDoc.PresentationRenderer.ConvertToImage(chart, image);
            //Closes the presentation
            pptxDoc.Close();
            image.Close();
            inputStream.Close();
        }
    }
}
