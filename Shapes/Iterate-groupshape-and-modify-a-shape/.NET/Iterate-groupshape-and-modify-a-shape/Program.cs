using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first group shape of the slide.
IGroupShape groupShape = slide.GroupShapes[0];
//Create an instance to hold shape collection.
IShapes shapes = groupShape.Shapes;
//Iterate the shape collection to remove the picture in a group shape.
foreach (IShape shape in shapes)
{
	if (shape.SlideItemType == SlideItemType.Picture)
	{
		shapes.Remove(shape);
		break;
	}
}
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);