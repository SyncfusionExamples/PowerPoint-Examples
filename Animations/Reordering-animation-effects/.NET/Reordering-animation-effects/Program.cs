using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation from the file system
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the slide
ISlide slide = pptxDoc.Slides[0];
//Retrieve the shape
IShape shape = slide.Shapes[0] as IShape;
//Access the animation sequence
ISequence sequence = slide.Timeline.MainSequence;
//Get the animation effects of the shape
IEffect[] shapeAnimationEffects = sequence.GetEffectsByShape(shape);
//Get the second animation effect of the shape
IEffect effect = shapeAnimationEffects[1];
//Remove the animation effect from the sequence
sequence.Remove(effect);
//Insert the removed animation effect as first
sequence.Insert(0, effect);
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));