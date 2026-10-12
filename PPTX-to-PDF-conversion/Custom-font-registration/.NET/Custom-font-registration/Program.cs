using Syncfusion.Drawing.Fonts;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Custom_font_registration
{
    class Program
    {
        public static void Main(string[] args)
        {
            FontManager.RegisterFonts(Path.GetFullPath(@"Fonts"));
            List<string> registerFont = FontManager.RegisteredFontNames;
            PPTXtoPDF();
            FontManager.ClearRegisteredFonts(true);
        }

        private static void PPTXtoPDF()
        {
            IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
            //Converts the PowerPoint Presentation into PDF document
            PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation);
            //Saves the PDF document
            pdfDocument.Save(Path.GetFullPath(@"Output/Output.pdf"));
            //Closes the PDF document
            pdfDocument.Close(true);
            //Closes the Presentation
            presentation.Close();
        }

    }
}
