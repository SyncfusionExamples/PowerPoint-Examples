using Syncfusion.Presentation;
using System;
using System.IO;
using System.Threading.Tasks;

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
                //Open and save multiple presentations, one PPT on each thread.
                OpenAndSavePresentationDocument(count);
                Console.WriteLine("Task {0} is done", count);
            });
        }
        //Open and save a Powerpoint  document using multi-threading.
        static void OpenAndSavePresentationDocument(int count)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.pptx"), FileMode.Open, FileAccess.Read))
            {
                // Load the PowerPoint presentation.
                using (IPresentation presentation = Presentation.Open(inputStream))
                {
                    // Save the presentation in the desired format.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + count + ".pptx"), FileMode.Create, FileAccess.Write))
                    {
                        presentation.Save(outputFileStream);
                    }
                }
            }
        }
    }

}
