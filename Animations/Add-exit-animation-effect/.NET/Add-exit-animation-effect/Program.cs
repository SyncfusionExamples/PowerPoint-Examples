using Syncfusion.Presentation;

//Create an instance for PowerPoint
IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Access the animation sequence to create effects
ISequence sequence = slide.Timeline.MainSequence;
//Add random bars effect to the shape
IEffect effect = sequence.AddEffect(cubeShape, EffectType.RandomBars, EffectSubtype.None, EffectTriggerType.OnClick);
//Change the preset class type of the effect from default entrance to exit
effect.PresetClassType = EffectPresetClassType.Exit;
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));