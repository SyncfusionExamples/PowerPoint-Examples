using Syncfusion.Presentation;
using Syncfusion.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;
using Syncfusion.PresentationRenderer;

namespace Multithreading_using_parallel_process
{
    class MultiThreading
    {
        static void Main(string[] args)
        {
            //Indicates the number of threads to be create.
            int limit = 5;
            Console.WriteLine("Parallel For Loop");
            Parallel.For(0, limit, count =>
            {
                Console.WriteLine("Task {0} started", count);
                //Create multiple presentations, one PPT on each thread.
                ConvertPowerPointToPDF(count);
                Console.WriteLine("Task {0} is done", count);
            });
        }
        //Convert a Powerpoint document to PDF using multi-threading.
        static void ConvertPowerPointToPDF(int count)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.pptx"), FileMode.Open, FileAccess.Read))
            {
                // Load the PowerPoint presentation.
                using (IPresentation presentation = Presentation.Open(inputStream))
                {
                    // Convert Power point document to PDF.
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation))
                    {
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + count + ".pdf"), FileMode.Create, FileAccess.Write))
                        {
                            // Save the PDF document
                            pdfDocument.Save(outputFileStream);
                        }
                    }
                }
            }
        }
    }
}
