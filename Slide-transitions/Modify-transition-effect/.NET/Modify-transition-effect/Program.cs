using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of a PowerPoint file.
ISlide slide = pptxDoc.Slides[0];
//Modify the transition effect applied to the slide.
slide.SlideTransition.TransitionEffect = TransitionEffect.Cover;
//Set the transition subtype.
slide.SlideTransition.TransitionEffectOption = TransitionEffectOption.Right;
//Save the PowerPoint Presentation as file
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));