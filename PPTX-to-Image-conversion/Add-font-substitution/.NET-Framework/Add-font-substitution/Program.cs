using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;

namespace Add_font_substitution
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Initialize 'ChartToImageConverter' to convert charts in the slides, and this is optional.
                pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                // Initializes the 'SubstituteFont' event to set the replacement font.
                pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                //Converts the first slide into image.
                Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
                //Saves the image as file.
                image.Save("../../slide1.png");
                //Disposes the image.
                image.Dispose();
            }
        }
        /// <summary>
        /// Sets the alternate font when a specified font is unavailable in the production environment.
        /// </summary>
        /// <param name="sender">FontSettings type of the Presentation in which the specified font is used but unavailable in production environment. </param>
        /// <param name="args">Retrieves the unavailable font name and receives the substitute font name for conversion. </param>
        private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            if (args.OriginalFontName == "Arial Unicode MS")
                args.AlternateFontName = "Arial";
            else
                args.AlternateFontName = "Times New Roman";
        }
    }
}
