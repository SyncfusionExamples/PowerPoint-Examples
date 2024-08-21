using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide of blank layout type.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add one more slide.
ISlide slide2 = pptxDoc.Slides.Add();
//Add normal shape to the first slide.
IShape shape = slide1.Shapes.AddShape(AutoShapeType.Rectangle, 100, 20, 200, 100);
//Set the target slide index (index is valid from 0 to slides count – 1) as hyperlink.
IHyperLink hyperLink = shape.SetHyperlink("1");
//Get the target slide of the hyperlink.
ISlide targetSlide = hyperLink.TargetSlide;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);