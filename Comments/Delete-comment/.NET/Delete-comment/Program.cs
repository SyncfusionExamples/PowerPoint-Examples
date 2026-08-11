using Syncfusion.Presentation;

//Open a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Get a comment from the slide
IComment comment = slide.Comments[0];
//Remove the comment from the slide
slide.Comments.Remove(comment);
//Save the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/DeleteComment.pptx"));
//Close the Presentation
pptxDoc.Close();