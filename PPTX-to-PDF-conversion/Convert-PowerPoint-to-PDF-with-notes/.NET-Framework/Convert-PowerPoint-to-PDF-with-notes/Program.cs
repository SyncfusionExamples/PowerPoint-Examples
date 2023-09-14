using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace Convert_PowerPoint_to_PDF_with_notes
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Create an instance of ChartToImageConverter and assigns it to ChartToImageConverter property of Presentation
                pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                //Enable the include hidden slides option in converter settings.
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings
                {
                    PublishOptions = PublishOptions.NotesPages
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