using Syncfusion.Drawing.Fonts;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Telemetry;

namespace Custom_font_registration
{
    class Program
    {
        public static void Main(string[] args)
        {
            FontManager.RegisterFonts(Path.GetFullPath(@"Fonts"));
            List<string> registerFont = FontManager.RegisteredFontNames;
            PPTXtoImage();
            FontManager.ClearRegisteredFonts(true);
        }

        private static void PPTXtoImage()
        {
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
            {
                //Initialize PresentationRenderer.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert the PowerPoint presentation as image streams.
                Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Png);
                //Save the image streams to file.
                for (int i = 0; i < images.Length; i++)
                {
                    using (Stream stream = images[i])
                    {
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/PPTXtoImage" + i + ".png")))
                        {
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }

    }
}
