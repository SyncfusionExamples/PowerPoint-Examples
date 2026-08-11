using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation sourcePresentation = Presentation.Open(Path.GetFullPath(@"Data/SourcePresentation.pptx"));
//Open the destination Presentation.
using IPresentation destinationPresentation = Presentation.Open(Path.GetFullPath(@"Data/DestinationPresentation.pptx"));
//Clone the first slide of the source Presentation.
ISlide clonedSlide = sourcePresentation.Slides[0].Clone();
//Merge the cloned slide to the destination Presentation with paste option - Destination Theme.
destinationPresentation.Slides.Add(clonedSlide, PasteOptions.SourceFormatting);
destinationPresentation.Save(Path.GetFullPath(@"Output/Result.pptx"));