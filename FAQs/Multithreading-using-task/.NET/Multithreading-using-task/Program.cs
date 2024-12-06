using Syncfusion.Presentation;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Multithreading_using_task
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
                tasks[i] = Task.Run(() => OpenAndSavePresentationDocument());
            }
            //Ensure all tasks complete by waiting on each task.
            await Task.WhenAll(tasks);
        }

        //Open and save a Powerpoint document using multi-threading.
        static void OpenAndSavePresentationDocument()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.pptx"), FileMode.Open, FileAccess.Read))
            {
                // Load the PowerPoint presentation.
                using (IPresentation presentation = Presentation.Open(inputStream))
                {
                    // Save the presentation in the desired format.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".pptx"), FileMode.Create, FileAccess.Write))
                    {
                        presentation.Save(outputFileStream);
                    }
                }
            }
        }
    }
}
