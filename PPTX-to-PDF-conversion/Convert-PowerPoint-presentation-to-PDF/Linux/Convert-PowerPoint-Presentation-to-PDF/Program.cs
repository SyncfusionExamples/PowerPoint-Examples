using System;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.IO;
using static System.Collections.Specialized.BitVector32;


namespace Convert_PowerPoint_Presentation_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"../../../Data/Input.pptx")))
            {
                //Convert the PowerPoint presentation to PDF document.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    //Save the PDF document to the file system.
                    pdfDocument.Save("Sample.pdf");
                }
            }
        }
    }
}
