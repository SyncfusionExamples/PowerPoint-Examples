using Syncfusion.Presentation;
using System.IO;
using Syncfusion.OfficeChartToImageConverter;

namespace Convert_chart_to_image
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the Presentation.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Initialize the ChartToImageConverter class; this is mandatory.
                    pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                    //Set the scaling mode for quality.
                    pptxDoc.ChartToImageConverter.ScalingMode = Syncfusion.OfficeChart.ScalingMode.Best;
                    //Get the first slide.
                    ISlide slide = pptxDoc.Slides[0];
                    //Get the chart in slide.
                    IPresentationChart chart = slide.Shapes[0] as IPresentationChart;
                    //Create a stream instance to store the image.
                    using (MemoryStream stream = new MemoryStream())
                    {
                        //Save the image to stream.
                        chart.SaveAsImage(stream);
                        //Save the stream to a file.
                        using (FileStream fileStream = File.Create(Path.GetFullPath(@"../../ChartImage.png"), (int)stream.Length))
                            fileStream.Write(stream.ToArray(), 0, stream.ToArray().Length);
                    }
                }
            }
        }
    }
}