using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Remove the specified slide from the Presentation.
pptxDoc.Slides.Remove(slide);
// Remove the slide from the specified index.
pptxDoc.Slides.RemoveAt(1);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));