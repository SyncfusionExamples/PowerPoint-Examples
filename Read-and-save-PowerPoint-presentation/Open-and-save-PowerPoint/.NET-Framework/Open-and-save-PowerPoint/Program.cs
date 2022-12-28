using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Open_and_save_PowerPoint
{
    internal class Program
    {
        static void Main()
        {
            //Loads or open an PowerPoint Presentation
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads or open an PowerPoint Presentation
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Gets the first slide from the PowerPoint presentation
                    ISlide slide = pptxDoc.Slides[0];
                    //Gets the first shape of the slide
                    IShape shape = slide.Shapes[0] as IShape;
                    //Change the text of the shape
                    if (shape.TextBody.Text == "Company History")
                        shape.TextBody.Text = "Company Profile";
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        pptxDoc.Save(outputStream);
                    } 
                }
            }
        }
    }
}
