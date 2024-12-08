using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Access the animation sequence to create effects.
ISequence sequence = slide.Timeline.MainSequence;
//Add bounce effect to the shape.
IEffect bounceEffect = sequence.AddEffect(cubeShape, EffectType.Bounce, EffectSubtype.None, EffectTriggerType.OnClick);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);