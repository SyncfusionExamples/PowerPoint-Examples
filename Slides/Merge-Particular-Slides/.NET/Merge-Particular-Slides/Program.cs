using Syncfusion.Presentation;

MergeParticularSlides(0, 2);


void MergeParticularSlides(int slideIndex, int count)
{
    //Load or open an existing source PowerPoint Presentation.
    using FileStream sourcePresentationStream = new(Path.GetFullPath(@"../../../Data/SourcePresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using IPresentation sourcePresentation = Presentation.Open(sourcePresentationStream);

    //Open the existing destination Presentation.
    using FileStream destinationPresentationStream = new(Path.GetFullPath(@"../../../Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using IPresentation destinationPresentation = Presentation.Open(destinationPresentationStream);

    //Iterate each slide in source and merge to presentation.
    for (int i = slideIndex; i < slideIndex + count; i++)
    {
        //Clone the slide of the source Presentation.
        ISlide clonedSlide = sourcePresentation.Slides[i].Clone();
        //Merge the cloned slide to the destination Presentation with paste option - Destination Theme.
        destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);
    }

    //Save the PowerPoint presentation.
    using FileStream outputStream = new(Path.GetFullPath(@"../../../Data/MergeParticularSlides.pptx"), FileMode.Create, FileAccess.ReadWrite);
    destinationPresentation.Save(outputStream);
}