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
IEffect wheelEffect = sequence[0] as IEffect;
//Change the wheel animation effect sub type from 2 spoke to 4 spoke
wheelEffect.Subtype = EffectSubtype.Wheel4;
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));