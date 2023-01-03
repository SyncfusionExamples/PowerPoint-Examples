using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Iterate through shapes in a slide and sets title.
foreach (IShape shape in pptxDoc.Slides[0].Shapes)
{
	if (shape is IPicture)
		shape.Title = "Picture";
	else if (shape is IShape)
		shape.Title = "AutoShape";
}
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);