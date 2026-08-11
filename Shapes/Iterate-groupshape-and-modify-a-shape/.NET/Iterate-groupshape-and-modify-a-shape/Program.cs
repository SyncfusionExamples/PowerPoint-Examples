using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
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
//Save the PowerPoint Presentation 
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));