using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a new LayoutSlide to the PowerPoint file.
ILayoutSlide layoutSlide = pptxDoc.Masters[0].LayoutSlides.Add(SlideLayoutType.Blank, "CustomLayout");
//Add a shape to the LayoutSlide.
IShape shape = layoutSlide.Shapes.AddShape(AutoShapeType.Diamond, 30, 20, 400, 300);
//Change the background color for LayoutSlide.
layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(78, 89, 90);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);