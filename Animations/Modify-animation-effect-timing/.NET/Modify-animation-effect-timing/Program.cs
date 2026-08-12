using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation from the file system
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape
IShape shape = slide.Shapes[0] as IShape;
//Access the animation main sequence to modify the effects
ISequence sequence = slide.Timeline.MainSequence;
//Get the required animation effect from the slide            
IEffect timingEffect = sequence[0] as IEffect;
//Increase the duration of the animation effect (value is in seconds)
timingEffect.Behaviors[0].Timing.Duration = 5;
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));