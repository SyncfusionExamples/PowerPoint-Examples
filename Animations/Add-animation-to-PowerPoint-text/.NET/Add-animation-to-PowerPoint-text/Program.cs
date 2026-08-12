using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation from the file system
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape
IShape shape = slide.Shapes[0] as IShape;
//Access the animation sequence to create effects
ISequence sequence = slide.Timeline.MainSequence;
//Add swivel effect with vertical subtype to the shape, build type is used to represent the animate level of the paragraph
IEffect swivelEffect = sequence.AddEffect(shape, EffectType.Swivel, EffectSubtype.Vertical, EffectTriggerType.OnClick, BuildType.ByLevelParagraphs1);
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));