
using Syncfusion.Presentation;
using System;
using System.ComponentModel;

namespace First_slide_number
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing PowerPoint Presentation.
            using (FileStream inputStream = new FileStream("../../../Data/Input.pptx", FileMode.Open))
            {
                using (IPresentation pptxDoc = Presentation.Open("../../../Data/Input.pptx"))
                {
                    //Get the FirstSlideNumber of Presentation.
                    int firstSlideNumber = pptxDoc.FirstSlideNumber;

                    //Modify the value for FirstSlideNumber.
                    pptxDoc.FirstSlideNumber = 10;

                    //Save the PowerPoint Presentation as stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create))
                    {
                        pptxDoc.Save(outputStream);
                    }                       
                }
            }
            // Open the PowerPoint presentation located at the specified path using the default associated program.
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Result.pptx"))
            {
                UseShellExecute = true
            };
            process.Start();
        }
    }
}
