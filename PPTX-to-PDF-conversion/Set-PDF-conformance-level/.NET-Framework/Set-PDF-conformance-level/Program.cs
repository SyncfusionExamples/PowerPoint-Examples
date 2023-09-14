using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace Set_PDF_conformance_level
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Initialize the conversion settings
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
                //Set the Pdf conformance level to A1B
                pdfConverterSettings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
                //Convert the documents by passing the settings as parameter.
                PdfDocument pdfDoc = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings);
                //Save the converted PDF document.
                pdfDoc.Save("../../PPTXToPDF.pdf");
            }
        }
    }
}