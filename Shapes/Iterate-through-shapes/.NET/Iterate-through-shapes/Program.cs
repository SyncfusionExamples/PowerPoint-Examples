using Syncfusion.Presentation;

//Loads or opens a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Iterates through the shapes in a slide and sets the title
foreach (IShape shape in pptxDoc.Slides[0].Shapes)
{
    if (shape is IPicture)
        shape.Title = "Picture";
    else if (shape is IShape)
        shape.Title = "AutoShape";
}
//Saves the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
//Closes the Presentation
pptxDoc.Close();