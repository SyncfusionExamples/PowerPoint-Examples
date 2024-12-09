using Syncfusion.Presentation;
using System;
using System.IO;
using System.Threading.Tasks;

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
                tasks[i] = Task.Run(() => OpenAndSavePresentation());
            }
            //Ensure all tasks complete by waiting on each task.
            await Task.WhenAll(tasks);
        }

        //Open and save a Powerpoint presentation using multi-threading.
        static void OpenAndSavePresentation()
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Load an existing PowerPoint presentation.
                using (IPresentation presentation = Presentation.Open(inputStream))
                {
                    //Add a slide of TitleAndContent type.
                    ISlide slide = presentation.Slides.Add(SlideLayoutType.TitleAndContent);
                    //Save the presentation in the desired format.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".pptx"), FileMode.Create, FileAccess.Write))
                    {
                        presentation.Save(outputFileStream);
                    }
                }
            }
        }
    }
}
