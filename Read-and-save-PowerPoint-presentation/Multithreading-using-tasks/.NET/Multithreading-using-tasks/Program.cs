using Syncfusion.Presentation;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Multithreading_using_tasks
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
            //Open an existing PowerPoint presentation.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Input.pptx")))
            {
                //Add a slide of TitleAndContent type.
                ISlide slide = presentation.Slides.Add(SlideLayoutType.TitleAndContent);
                //Save the presentation in the desired format.
                presentation.Save(Path.GetFullPath(@"Output/Output" + Guid.NewGuid().ToString() + ".pptx"));
            }
        }
    }
}
