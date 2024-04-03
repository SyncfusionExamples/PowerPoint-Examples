using Syncfusion.Presentation;


List<string> sourceFiles = new List<string> { "Product overview.pptx", "About.pptx"};

//Open the existing destination Presentation.
using FileStream destinationPresentationStream = new(Path.GetFullPath(@"../../../Data/Adventure.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using IPresentation destinationPresentation = Presentation.Open(destinationPresentationStream);

string path = Path.GetFullPath(@"../../../Data/");
foreach (string file in sourceFiles)
{
    //Load or open an existing source PowerPoint Presentation.
    using FileStream sourcePresentationStream = new( path + file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using IPresentation sourcePresentation = Presentation.Open(sourcePresentationStream);
    //Iterate each slide in source and merge to presentation.
    for (int i = 0; i < sourcePresentation.Slides.Count; i++)
    {
        //Clone the slide of the source Presentation.
        ISlide clonedSlide = sourcePresentation.Slides[i].Clone();
        //Merge the cloned slide to the destination Presentation with paste option - Destination Theme.
        destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);
    }
}

//Save the PowerPoint presentation.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Data/MergeMultipleFiles.pptx"), FileMode.Create, FileAccess.ReadWrite);
destinationPresentation.Save(outputStream);