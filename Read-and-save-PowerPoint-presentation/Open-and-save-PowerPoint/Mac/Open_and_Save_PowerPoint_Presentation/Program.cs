using Syncfusion.Presentation;

namespace Open_and_Save_PowerPoint_Presentation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing PowerPoint presentation
            IPresentation pptxDoc = Presentation.Open(new FileStream(Path.GetFullPath(@"../../../Data/Input.pptx"), FileMode.Open));
            //Gets the first slide from the PowerPoint presentation
            ISlide slide = pptxDoc.Slides[0];
            //Gets the first shape of the slide
            IShape shape = slide.Shapes[0] as IShape;
            //Change the text of the shape
            if (shape.TextBody.Text == "Company History")
                shape.TextBody.Text = "Company Profile";
            //Save the PowerPoint presentation as stream
            FileStream outputStream = new FileStream("Output.pptx", FileMode.Create);
            pptxDoc.Save(outputStream);
            outputStream.Position = 0;
            outputStream.Flush();
            outputStream.Dispose();
            //Close the PowerPoint presentation
            pptxDoc.Close();
        }
    }
}

           
