using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Iterate the slide.
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape.
IShape shape = slide.Shapes[0] as IShape;
//Iterate the sequence.
ISequence sequence = slide.Timeline.MainSequence;
//To Remove the animation effects from the shape.
//Get the animation effects of the particular shape.
IEffect[] animationEffects = sequence.GetEffectsByShape(shape);
//Remove the animation effect from the main sequence.
foreach (IEffect effect in animationEffects)
{
    sequence.Remove(effect);
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);