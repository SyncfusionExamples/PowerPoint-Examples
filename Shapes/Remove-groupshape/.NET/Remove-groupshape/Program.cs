using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first group shape of the slide.
IGroupShape groupShape = slide.GroupShapes[0];
//Remove the group shape from group shape collection.
slide.GroupShapes.Remove(groupShape);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));