using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace PowerPoint_to_PDF_with_hidden_slides
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Enable or disable including the hidden slides option in converter settings.
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings
                {
                    ShowHiddenSlides = true
                };
                //Convert the documents by passing the settings as parameter.
                using (PdfDocument pdfDoc = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings))
                {
                    //Save the converted PDF document.
                    pdfDoc.Save("../../PPTXToPDF.pdf");
                }
            }
        }
    }
}