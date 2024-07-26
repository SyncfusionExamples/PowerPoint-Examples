using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Get_missing_fonts_for_PDF_conversion
{
    class Program
    {
        // List to store names of fonts that are not installed
        static List<string> fonts = new List<string>();
        static void Main()
        {
            // Open the existing PowerPoint presentation file stream
            using (FileStream pptxStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                // Load the file stream into a PowerPoint presentation
                using (IPresentation pptxDoc = Presentation.Open(pptxStream))
                {
                    // Hook the font substitution event to detect missing fonts
                    pptxDoc.FontSettings.SubstituteFont += FontSettings_SubstituteFont;

                    // Convert PowerPoint presentation into PDF document
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        // Save the PDF document to output stream
                        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Data/Result.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                        {
                            pdfDocument.Save(outputStream);
                        }
                    }
                }
            }

            // Print the fonts that are not available in machine, but used in Word document.
            if (fonts.Count > 0)
            {
                Console.WriteLine("Fonts not available in environment:");
                foreach (string font in fonts)
                    Console.WriteLine(font);
            }
            else
            {
                Console.WriteLine("Fonts used in Word document are available in environment.");
            }
			Console.ReadKey();
        }

        // Event handler for font substitution event
        static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            // Add the original font name to the list if it's not already there
            if (!fonts.Contains(args.OriginalFontName))
                fonts.Add(args.OriginalFontName);
        }
    }

}
