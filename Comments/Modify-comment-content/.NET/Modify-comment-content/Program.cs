using Syncfusion.Presentation;

//Open a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Get the comment from the slide
IComment comment = slide.Comments[0] as IComment;
//Modify the comment text
comment.Text = "The comment text content is changed";
//Save the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/ModifyCommentText.pptx"));
//Close the Presentation
pptxDoc.Close();