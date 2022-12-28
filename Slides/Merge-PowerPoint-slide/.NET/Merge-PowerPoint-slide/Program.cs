using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream sourcePresentationStream = new(Path.GetFullPath(@"../../../Data/SourcePresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    using (FileStream destinationPresentationStream = new(Path.GetFullPath(@"../../../Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        //Open an existing PowerPoint presentation.
        using IPresentation sourcePresentation = Presentation.Open(sourcePresentationStream);
        //Opens the destination Presentation
        using IPresentation destinationPresentation = Presentation.Open(destinationPresentationStream);

        //Clones the first slide of the source Presentation
        ISlide clonedSlide = sourcePresentation.Slides[0].Clone();
        //Merges the cloned slide to the destination Presentation with paste option - Destination Theme
        destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);

        using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
        destinationPresentation.Save(outputStream);
    }   
}