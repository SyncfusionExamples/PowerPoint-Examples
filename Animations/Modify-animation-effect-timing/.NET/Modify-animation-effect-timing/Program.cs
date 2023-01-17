using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieves the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Access the animation main sequence to modify the effects.
ISequence sequence = slide.Timeline.MainSequence;
//Get the required animation effect from the slide.        
IEffect pathEffect = sequence[0] as IEffect;
//Increase the duration of the animation effect.
pathEffect.Behaviors[0].Timing.Duration = 5;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);