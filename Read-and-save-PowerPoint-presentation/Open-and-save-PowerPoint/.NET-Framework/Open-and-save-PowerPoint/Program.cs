using Syncfusion.Presentation;
using System.IO;

namespace Open_and_save_PowerPoint
{
    internal class Program
    {
        static void Main()
        {
            //Loads or open an PowerPoint Presentation
            IPresentation pptxDoc = Presentation.Open(@"../../Data/Template.pptx");
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
                //Closes the Presentation instance and free the memory consumed.
                pptxDoc.Close();
            }
        }
    }
}
