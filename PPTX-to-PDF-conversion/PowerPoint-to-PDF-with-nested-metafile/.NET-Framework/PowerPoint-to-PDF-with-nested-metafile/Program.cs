using System.Drawing;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using System;
using System.IO;
using System.Drawing.Imaging;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace PowerPoint_to_PDF_with_nested_metafile
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Create an instance of the ChartToImageConverter and assign it to the ChartToImageConverter property of the Presentation.
                pptxDoc.ChartToImageConverter = new ChartToImageConverter();
                //Initialize the conversion settings.
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
                //Set the RecreateNestedMetafile property to true to recreate the nested metafile automatically.          
                pdfConverterSettings.RecreateNestedMetafile = true;
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