using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.OfficeChart;

namespace Set_chart_quality
{
    internal class Program
    {
        static void Main()
        {
            //Load or open an PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx"))
			{
                //Create an instance of ChartToImageConverter and assigns it to ChartToImageConverter property of Presentation
                pptxDoc.ChartToImageConverter = new ChartToImageConverter
                {
                    //Sets the scaling mode of the chart to best. 
                    ScalingMode = ScalingMode.Best
                };
                //Convert the documents by passing the PDF conversion settings as parameter.
                using (PdfDocument pdfDoc = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Save the converted PDF document.
                    pdfDoc.Save("../../PPTXToPDF.pdf");
                }
            }
        }
    }
}