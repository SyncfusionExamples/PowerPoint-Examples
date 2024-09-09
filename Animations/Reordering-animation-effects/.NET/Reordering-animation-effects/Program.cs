using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Iterate the slide.
ISlide slide = pptxDoc.Slides[0];
//Iterate the shape.
IShape shape = slide.Shapes[0] as IShape;
//Iterate the sequence.
ISequence sequence = slide.Timeline.MainSequence;
//Get the animation effects of the shape.
IEffect[] shapeAnimationEffects = sequence.GetEffectsByShape(shape);
//Get the second animation effect of the shape.
IEffect effect = shapeAnimationEffects[1];
//Remove the animation effect from the sequence.
sequence.Remove(effect);
//Insert the removed animation effect as first.
sequence.Insert(0, effect);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);