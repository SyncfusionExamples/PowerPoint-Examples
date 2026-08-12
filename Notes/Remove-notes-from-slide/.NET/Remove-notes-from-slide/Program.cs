using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get instance of the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0] as ISlide;
//Remove Notes Slide from a corresponding slide.
slide.RemoveNotesSlide();
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));