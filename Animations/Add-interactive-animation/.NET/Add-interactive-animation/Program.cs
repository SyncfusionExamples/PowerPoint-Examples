using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 200, 100, 300, 300);
//Add a shape to act as button.
IShape buttonShape = slide.Shapes.AddShape(AutoShapeType.Oval, 130, 75, 50, 50);
//Create the interactive sequence to make the animation effects interactive by triggering with button click.
ISequence interactiveSequence = slide.Timeline.InteractiveSequences.Add(buttonShape);
//Add Fly effect with top subtype to animate the shape as fly from top.
IEffect bounceEffect = interactiveSequence.AddEffect(cubeShape, EffectType.Fly, EffectSubtype.Top, EffectTriggerType.OnClick);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);