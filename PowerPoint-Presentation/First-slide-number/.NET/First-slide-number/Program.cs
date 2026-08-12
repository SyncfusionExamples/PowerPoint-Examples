
using Syncfusion.Presentation;
using System;
using System.ComponentModel;

namespace First_slide_number
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing PowerPoint presentation
            IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Input.pptx"));
            //Gets the FirstSlideNumber of the presentation
            int firstSlideNumber = pptxDoc.FirstSlideNumber;
            //Modifies the value for the FirstSlideNumber
            pptxDoc.FirstSlideNumber = 10;
            //Saves the PowerPoint presentation
            pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
            //Closes the PowerPoint presentation
            pptxDoc.Close();
        }
    }
}
