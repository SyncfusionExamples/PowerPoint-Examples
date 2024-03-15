using Syncfusion.Office;
using Syncfusion.Presentation;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Drawing.Imaging;

namespace Extract_crop_image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open an existing PowerPoint Presentation.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../Data/Sample.pptx"), FileMode.Open, FileAccess.Read))
            {
                using (IPresentation presentation = Presentation.Open(Path.GetFullPath(@"../../Data/Sample.pptx")))
                {
                    //Get the first slide.
                    ISlide slide = presentation.Slides[0];
                    //Get the picture.
                    IPicture picture = slide.Pictures[0];
                    //Get cropped image stream.
                    using (MemoryStream stream = GetCroppedPictureData(new MemoryStream(picture.ImageData), picture.Crop))
                    {
                        //Save the image as stream.
                        using (FileStream fileStreamOutput = new FileStream(Path.GetFullPath(@"../../../Output.png"), FileMode.Create))
                        {
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }                    
        }


        /// <summary>
        /// Get the cropped picture data stream.
        /// </summary>
        /// <param name="picture">The picture to be cropped.</param>
        /// <returns>MemoryStream containing the cropped picture data.</returns>
        private static MemoryStream GetCroppedPictureData(MemoryStream imageStream, ICrop crop)
        {
            //Get the crop properties.
            float imageWidth = crop.ContainerWidth;
            float imageHeight = crop.ContainerHeight;
            float offsetX = crop.OffsetX;
            float offsetY = crop.OffsetY;
            float rectWidth = crop.Width;
            float rectHeight = crop.Height;
            float rectX = (imageWidth - rectWidth) / 2;
            float rectY = (imageHeight - rectHeight) / 2;
            //Creating memory stream to return the cropped image stream.
            MemoryStream memoryStream = new MemoryStream();

            //Create the image with respect to given image size.
            using (Image outputImage = new Bitmap(PointToPixel(imageWidth), PointToPixel(imageHeight), PixelFormat.Format32bppPArgb))
            {
                //Draw the given image with in the image graphics.
                using (Graphics graphics = Graphics.FromImage(outputImage))
                {
                    //Set the resolution for the image graphics
                    (outputImage as Bitmap).SetResolution(graphics.DpiX, graphics.DpiY);
                    //Set the background color of the image as transparent.
                    graphics.Clear(Color.Transparent);
                    //Set the graphics page unit is Point.
                    graphics.PageUnit = GraphicsUnit.Point;

                    //Create image for the given image data.
                    using (Image inputImage = new Bitmap(imageStream))
                    {
                        // Translate the graphics and apply the Crop.OffsetX and Crop.OffsetY values.
                        graphics.TranslateTransform(offsetX, offsetY);

                        // Draw the cropped image at the specified picture position.
                        RectangleF srcRect = new RectangleF(rectX, rectY, rectWidth, rectHeight);
                        graphics.DrawImage(inputImage, srcRect);
                    }
                }
                //Save the image as stream.
                outputImage.Save(memoryStream, ImageFormat.Png);
                memoryStream.Position = 0;
            }
            return memoryStream;
        }

        /// <summary>
        /// Convert a value from point to pixel.
        /// </summary>
        /// <param name="value">The value in point which needs to be converted to pixel.</param>
        /// <returns>The value converted to pixel.</returns>
        internal static int PointToPixel(double value)
        {
            return Convert.ToInt32((value * 96) / 72);
        }
    }
}
