using Syncfusion.Presentation;

//Open a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Get the comment in the slide
IComment comment = slide.Comments[0] as IComment;
//Add reply to the comment
slide.Comments.Add("Author2", "A2", "Yes, we can change the font size to 20", DateTime.Now, comment);
//Save the Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/ReplyComment.pptx"));
//Close the Presentation
pptxDoc.Close();