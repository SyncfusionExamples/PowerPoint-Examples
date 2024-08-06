using Syncfusion.Presentation;

namespace Set_SlideSize_to_destination_slides
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Open first source PowerPoint presentation with the stream.
            using (Presentation sourcePresentation1 = Presentation.Open(Path.GetFullPath(@"../../../Data/Presentation1.pptx")) as Presentation) 
            {
                //Open second source PowerPoint presentation with the stream.
                using (Presentation sourcePresentation2 = Presentation.Open(Path.GetFullPath(@"../../../Data/Presentation2.pptx")) as Presentation) 
                {
                    //Create a destination PowerPoint presentation.
                    using (Presentation destinationPresentation = Presentation.Create() as Presentation)
                    {
                        //Set the first source presentation slide size to the destination presentation slide size.
                        destinationPresentation.SlideSize.Type = sourcePresentation1.SlideSize.Type;
                        destinationPresentation.SlideSize.SlideOrientation = sourcePresentation1.SlideSize.SlideOrientation;
                        destinationPresentation.SlideSize.Width = sourcePresentation1.SlideSize.Width;
                        destinationPresentation.SlideSize.Height = sourcePresentation1.SlideSize.Height;

                        //Clone and merge the PowerPoint slides with PasteOptions.SourceFormatting.
                        foreach (ISlide slide in sourcePresentation1.Slides)
                            destinationPresentation.Slides.Add(slide, PasteOptions.SourceFormatting);
                        //Clone and merge the PowerPoint slides with PasteOptions.SourceFormatting.
                        foreach (ISlide slide in sourcePresentation2.Slides)
                            destinationPresentation.Slides.Add(slide, PasteOptions.SourceFormatting);

                        //Save the PowerPoint Presentation.
                        destinationPresentation.Save("Output.pptx");
                    }
                }            
            }
        }
    }
}
