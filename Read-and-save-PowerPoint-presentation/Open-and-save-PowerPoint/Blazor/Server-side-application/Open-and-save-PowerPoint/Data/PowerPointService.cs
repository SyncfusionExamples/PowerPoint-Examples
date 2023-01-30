using Syncfusion.Presentation;
using System.IO;

namespace Open_and_save_PowerPoint.Data
{
    public class PowerPointService
    {
        public MemoryStream GeneratePresentation()
        {
            using (FileStream inputStream = new FileStream(@"wwwroot/sample-data/Template.pptx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load or open an PowerPoint Presentation.
                using (IPresentation pptxDoc = Presentation.Open(inputStream))
                {
                    //Get the first slide from the PowerPoint presentation.
                    ISlide slide = pptxDoc.Slides[0];
                    //Get the first shape of the slide.
                    IShape shape = slide.Shapes[0] as IShape;
                    //Change the text of the shape.
                    if (shape.TextBody.Text == "Company History")
                        shape.TextBody.Text = "Company Profile";
                    //Saves the PowerPoint document to MemoryStream.
                    using (MemoryStream stream = new MemoryStream())
                    {
                        pptxDoc.Save(stream);
                        stream.Position = 0;
                        return stream;
                    }
                }
            }
        }
    }
}
