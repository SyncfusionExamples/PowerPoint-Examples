using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;

namespace Change_resolution_of_converted_image
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Declare variables to hold custom width and height.
                int customWidth = 1500;
                int customHeight = 1000;
                //Convert the slide as image and returns the image stream.
                Stream stream = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageFormat.Emf);
                //Create a bitmap of specific width and height.
                Bitmap bitmap = new Bitmap(customWidth, customHeight, PixelFormat.Format32bppPArgb);
                //Get the graphics from image.
                Graphics graphics = Graphics.FromImage(bitmap);
                //Set the resolution.
                bitmap.SetResolution(graphics.DpiX, graphics.DpiY);
                //Recreate the image in custom size.
                graphics.DrawImage(System.Drawing.Image.FromStream(stream), new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                //Save the image as bitmap .
                bitmap.Save("../../ImageOutput" + Guid.NewGuid().ToString() + ".jpeg");
            }
        }
    }
}
