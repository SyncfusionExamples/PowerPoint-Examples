using Syncfusion.Presentation;

namespace Remove_hidden_slides_in_powerpoint
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Loads or open an PowerPoint Presentation.
            using (FileStream inputStream = new FileStream(@"../../../Data/Template.pptx", FileMode.Open))
            {
                IPresentation pptxDoc = Presentation.Open(inputStream);

                //Remove every hidden slide.
                RemoveHiddenSlides(pptxDoc);

                //Save the PowerPoint Presentation as stream
                using (FileStream outputStream = new FileStream(@"../../../Output.pptx", FileMode.Create))
                {
                    pptxDoc.Save(outputStream);
                    //Closes the Presentation instance
                    pptxDoc.Close();
                }
            }
        }

        /// <summary>
        /// Removes the hidden slide
        /// </summary>
        static void RemoveHiddenSlides(IPresentation pptxDoc)
        {
            for (int i = 0; i < pptxDoc.Slides.Count; i++)
            {
                if (!pptxDoc.Slides[i].Visible)
                {
                    pptxDoc.Slides.Remove(pptxDoc.Slides[i]);
                    i--;
                }
            }
        }
    }
}