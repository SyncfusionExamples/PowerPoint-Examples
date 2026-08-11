using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Creates an instance for PowerPoint
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds normal shape to slide
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 200, 0, 300, 300);
//Access the animation sequence to create effects
ISequence sequence = slide.Timeline.MainSequence;
//Add user path effect to the shape
IEffect pathEffect = sequence.AddEffect(cubeShape, EffectType.PathUser, EffectSubtype.None, EffectTriggerType.OnClick);
//Add commands to the empty path for moving
IMotionEffect motionBehavior = ((IMotionEffect)pathEffect.Behaviors[0]);
PointF[] points = new PointF[1];
//Add the move command to move the position of the shape
points[0] = new PointF(0, 0);
motionBehavior.Path.Add(MotionCommandPathType.MoveTo, points, MotionPathPointsType.Auto, false);
//Add the line command to move the shape in straight line
points[0] = new PointF(0, 0.25f);
motionBehavior.Path.Add(MotionCommandPathType.LineTo, points, MotionPathPointsType.Auto, false);
//Add the end command to finish the path animation
motionBehavior.Path.Add(MotionCommandPathType.End, null, MotionPathPointsType.Auto, false);
//Save the PowerPoint Presentation to the file system
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));