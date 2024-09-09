using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first group shape of the slide.
IGroupShape groupShape = slide.GroupShapes[0];
//Remove the group shape from group shape collection.
slide.GroupShapes.Remove(groupShape);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);