using Syncfusion.Presentation;
using Syncfusion.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;
using Syncfusion.PresentationRenderer;

namespace Multithreaded_using_tasks
{
    class MultiThreading
    {
        //Indicates the number of threads to be create.
        private const int TaskCount = 1000;
        public static async Task Main()
        {
            //Create an array of tasks based on the TaskCount.
            Task[] tasks = new Task[TaskCount];
            for (int i = 0; i < TaskCount; i++)
            {
                tasks[i] = Task.Run(() => ConvertPowerPointToPDF());
            }
            //Ensure all tasks complete by waiting on each task.
            await Task.WhenAll(tasks);
        }

        //Convert a Powerpoint presentation to PDF using multi-threading.
        static void ConvertPowerPointToPDF()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Load an existing PowerPoint presentation.
                using (IPresentation presentation = Presentation.Open(inputStream))
                {
                    //Convert PowerPoint presentation to PDF.
                    using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(presentation))
                    {
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".pdf"), FileMode.Create, FileAccess.Write))
                        {
                            //Save the PDF document.
                            pdfDocument.Save(outputFileStream);
                        }
                    }
                }
            }
        }
    }
}
