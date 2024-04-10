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
            //Open the file as Stream
            using (FileStream fileStreamInput = new FileStream(Path.GetFullPath(@"../../../Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation with loaded stream.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    //Convert the PowerPoint document to PDF document.
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                    {
                        //Save the converted PDF document to MemoryStream.
                        MemoryStream pdfStream = new MemoryStream();
                        pdfDocument.Save(pdfStream);
                        pdfStream.Position = 0;
                        //Save the stream as file.
                        using (FileStream fileStreamOutput = File.Create("Sample.pdf"))
                        {
                            pdfStream.CopyTo(fileStreamOutput);
                        }
                    }                         
                }
            }
        }
    }
}
