using Syncfusion.Presentation;

//Create a PowerPoint Presentation
IPresentation pptxDoc = Presentation.Create();
//Add a slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a comment to the slide
slide.Comments.Add(10, 10, "Author1", "A1", "Can we change the font size to 20?", DateTime.Now);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Comment.pptx"));
//Close the Presentation
pptxDoc.Close();