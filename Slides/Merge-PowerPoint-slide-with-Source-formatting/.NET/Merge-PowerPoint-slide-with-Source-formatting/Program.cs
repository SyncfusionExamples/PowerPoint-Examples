using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream sourcePresentationStream = new(Path.GetFullPath(@"Data/SourcePresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using FileStream destinationPresentationStream = new(Path.GetFullPath(@"Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation sourcePresentation = Presentation.Open(sourcePresentationStream);
//Open the destination Presentation.
using IPresentation destinationPresentation = Presentation.Open(destinationPresentationStream);
//Clone the first slide of the source Presentation.
ISlide clonedSlide = sourcePresentation.Slides[0].Clone();
//Merge the cloned slide to the destination Presentation with paste option - Destination Theme.
destinationPresentation.Slides.Add(clonedSlide, PasteOptions.SourceFormatting);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
destinationPresentation.Save(outputStream);