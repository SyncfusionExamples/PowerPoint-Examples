using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation from the file system
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the slide
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape
IShape shape = slide.Shapes[0] as IShape;
//Access the animation sequence
ISequence sequence = slide.Timeline.MainSequence;
//Get the animation effects of the particular shape
IEffect[] animationEffects = sequence.GetEffectsByShape(shape);
//Remove the animation effects from the main sequence
foreach (IEffect effect in animationEffects)
{
    sequence.Remove(effect);
}
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));