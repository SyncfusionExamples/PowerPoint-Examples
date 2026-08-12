using Syncfusion.Presentation;

//Open a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Remove the comment at index 1 from the slide
slide.Comments.RemoveAt(1);
//Save the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/DeleteReplyComment.pptx"));
//Close the Presentation
pptxDoc.Close();