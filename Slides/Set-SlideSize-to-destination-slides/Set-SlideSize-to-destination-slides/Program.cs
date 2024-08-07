using Syncfusion.Presentation;

namespace Set_SlideSize_to_destination_slides
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Open the existing source PowerPoint presentation.
            using (Presentation? sourcePresentation = Presentation.Open(Path.GetFullPath(@"../../../Data/Template.pptx")) as Presentation) 
            {   //Create a destination PowerPoint presentation.
                using (Presentation? destinationPresentation = Presentation.Create() as Presentation)
                {
                    //Set the source presentation slide size to the destination presentation slide size.
                    destinationPresentation.SlideSize.Type = sourcePresentation.SlideSize.Type;
                    destinationPresentation.SlideSize.SlideOrientation = sourcePresentation.SlideSize.SlideOrientation;
                    destinationPresentation.SlideSize.Width = sourcePresentation.SlideSize.Width;
                    destinationPresentation.SlideSize.Height = sourcePresentation.SlideSize.Height;

                    //Clone and merge the PowerPoint slides with PasteOptions.SourceFormatting.
                    foreach (ISlide slide in sourcePresentation.Slides)
                        destinationPresentation.Slides.Add(slide, PasteOptions.SourceFormatting);

                    //Save the PowerPoint Presentation.
                    destinationPresentation.Save("Output.pptx");
                }          
            }
        }
    }
}
