using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Access the animation main sequence to modify the effects.
ISequence sequence = slide.Timeline.MainSequence;
//Get the animation effects of the particular shape.
IEffect[] animationEffects = sequence.GetEffectsByShape(shape);
//Iterate the animation effect to make the change.
IEffect animationEffect = animationEffects[0];
//Change the animation effect type from swivel to GrowAndTurn.
animationEffect.Type = EffectType.GrowAndTurn;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);