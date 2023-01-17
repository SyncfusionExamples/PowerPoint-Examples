using Syncfusion.Presentation;

//Create a PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a shape to the slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Set the transition effect type.
slide.SlideTransition.TransitionEffect = TransitionEffect.Checkerboard;
//Enable the transition time delay.
slide.SlideTransition.TriggerOnTimeDelay = true;
//Assign the value for the advance time delay in seconds.
slide.SlideTransition.TimeDelay = 5;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);