using System.Drawing;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Enable_text_shrink_on_overflow
{
    internal class Program
    {
        static void Main()
        {
            // Create a new PowerPoint file.
            using (IPresentation ppDoc = Presentation.Create())
            {
                //Add a slide to the PowerPoint file.
                ISlide slide = ppDoc.Slides.Add(SlideLayoutType.Blank);
                //Add a text box to the slide
                IShape textBox = slide.Shapes.AddTextBox(100, 100, 100, 100);
                //Add text to the text box. 
                textBox.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
                //Set the property to shrink text on overflow. 
                textBox.TextBody.FitTextOption = FitTextOption.ShrinkTextOnOverFlow;
                //Save the PowerPoint file
                ppDoc.Save(Path.GetFullPath(@"../../Result.pptx"));
            }
        }
    }
}
