using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Add_font_substitution
{
    internal class Program
    {
        static void Main()
        {
            //Load the PowerPoint presentation and convert to image
            using (IPresentation pptxDoc = Presentation.Open(@"../../../Data/Template.pptx"))
            {
                //Initialize the PresentationRenderer to perform image conversion.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                // Initializes the 'SubstituteFont' event to set the replacement font
                pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                //Convert PowerPoint slide to image as stream.
                using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                {
                    //Create the output image file stream
                    using (FileStream fileStreamOutput = File.Create("Output.jpg"))
                    {
                        //Copy the converted image stream into created output stream
                        stream.CopyTo(fileStreamOutput);
                    }
                }
            }
        }
        /// <summary>
        /// Sets the alternate font when a specified font is unavailable in the production environment
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