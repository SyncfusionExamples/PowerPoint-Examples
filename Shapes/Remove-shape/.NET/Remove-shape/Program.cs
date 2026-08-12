using Syncfusion.Presentation;

//Loads or opens a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieves the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Retrieves the first shape
IShape shape = slide.Shapes[0] as IShape;
//Removes the shape from the shape collection
slide.Shapes.Remove(shape);
//Saves the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
//Closes the Presentation
pptxDoc.Close();