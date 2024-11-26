using Syncfusion.Presentation;

//Open an existing source presentation.
using FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourcePresentation.pptx"), FileMode.Open);
using IPresentation sourcePresentation = Presentation.Open(sourceStream);
//Open an existing destination presentation.
using FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using IPresentation destinationPresentation = Presentation.Open(destinationStream);
//Iterate each slide in source and merge to presentation.
for (int i = 0; i < sourcePresentation.Slides.Count; i++)
{
    //Clone the slide of the source presentation.
    ISlide clonedSlide = sourcePresentation.Slides[i].Clone();
    //Merge the cloned slide into the destination presentation.
    destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);
}
//Save the PowerPoint presentation.
using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create);
destinationPresentation.Save(outputStream);
