using Syncfusion.Presentation;

//Open a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Get the comment from the slide
IComment comment = slide.Comments[0] as IComment;
//Modify the comment author name
comment.AuthorName = "NewAuthor";
//Save the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/ModifyCommentAuthor.pptx"));
//Close the Presentation
pptxDoc.Close();