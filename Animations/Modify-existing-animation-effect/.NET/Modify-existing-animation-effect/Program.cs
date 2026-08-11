using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation from the file system
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape
IShape shape = slide.Shapes[0] as IShape;
//Access the animation main sequence to modify the effects
ISequence sequence = slide.Timeline.MainSequence;
//Get the animation effects of the particular shape
IEffect[] animationEffects = sequence.GetEffectsByShape(shape);
//Iterate the animation effect to make the change
IEffect animationEffect = animationEffects[0];
//Change the animation effect type to GrowAndTurn
animationEffect.Type = EffectType.GrowAndTurn;
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
