
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System.IO;

namespace Convert_PowerPoint_presentation_to_PDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open a PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"../../Data/Template.pptx")))
            {
                //Convert the PowerPoint Presentation into PDF document.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Save the PDF document.
                    pdfDocument.Save(Path.GetFullPath(@"../../../Sample.pdf"));
                }
            }
            // Open the PowerPoint located at the specified path using the default associated program.
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Sample.pdf"))
            {
                UseShellExecute = true
            };
            process.Start();
        }
    }
}
