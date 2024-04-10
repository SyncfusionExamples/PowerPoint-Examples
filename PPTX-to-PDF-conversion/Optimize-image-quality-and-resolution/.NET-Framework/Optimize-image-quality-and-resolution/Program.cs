using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.OfficeChartToImageConverter;

namespace Optimize_image_quality_and_resolution
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
                    //Set the image resolution.
                    ImageResolution = 100,
                    //Set the image quality.
                    ImageQuality = 100
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