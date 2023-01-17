using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from the presentation.
ISlide slide = pptxDoc.Slides[0];
//Modify the transition effect applied to the slide.
slide.SlideTransition.TransitionEffect = TransitionEffect.Cover;
//Set the transition subtype.
slide.SlideTransition.TransitionEffectOption = TransitionEffectOption.Right;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);