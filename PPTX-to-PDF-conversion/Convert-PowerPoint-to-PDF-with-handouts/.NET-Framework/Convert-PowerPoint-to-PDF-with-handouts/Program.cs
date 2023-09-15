using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace Convert_PowerPoint_to_PDF_with_handouts
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Enable the handouts and number of pages per slide options in converter settings.
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings
                {
                    PublishOptions = PublishOptions.Handouts,
                    SlidesPerPage = SlidesPerPage.Nine
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