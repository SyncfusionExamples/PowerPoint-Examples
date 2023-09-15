using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;

namespace Convert_PowerPoint_presentation_to_Image
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Create instance of ChartToImageConverter.
                pptxDoc.ChartToImageConverter = new ChartToImageConverter
                {
                    //Set the scaling mode as best.
                    ScalingMode = Syncfusion.OfficeChart.ScalingMode.Best
                };
                //Convert entire Presentation to images.
                Image[] images = pptxDoc.RenderAsImages(Syncfusion.Drawing.ImageType.Metafile);
                int i = 1;
                //Save the image to file system.
                foreach (Image image in images)
                {
                    image.Save("../../Output" + i++ + ".png");
                    //Disposes the image.
                    image.Dispose();
                }
            }
        }
    }
}
