using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;

namespace Optimize_duplicate_images
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Enable the include hidden slides option in converter settings.
                PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings
                {
                    //Set the flag to optimize the identical images.
                    OptimizeIdenticalImages = true
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