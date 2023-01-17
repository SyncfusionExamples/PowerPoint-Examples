using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape.
IShape shape = slide.Shapes[1] as IShape;
//Access the animation main sequence to modify the effects.
ISequence sequence = slide.Timeline.MainSequence;
//Get the required animation effect from the slide.
IEffect wheelEffect = sequence[0] as IEffect;
//Change the wheel animation effect sub type from 2 spoke to 4 spoke.
wheelEffect.Subtype = EffectSubtype.Wheel4;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);