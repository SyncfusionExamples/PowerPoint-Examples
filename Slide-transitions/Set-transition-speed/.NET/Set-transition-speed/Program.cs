using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a shape to the slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Set the transition effect type.
slide.SlideTransition.TransitionEffect = TransitionEffect.Checkerboard;
//Set the transition effect speed enumeration. This will reduce the transition duration to 0.75 seconds from the default 2 second.
slide.SlideTransition.Speed = TransitionSpeed.Medium;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);