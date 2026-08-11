using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the first shape of the slide.
IShape shape = slide.Shapes[0] as IShape;
//Change the text of the shape.
if (shape.TextBody.Text == "Company History")
	shape.TextBody.Text = "Company Profile";
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
//Close the Presentation instance and free the memory consumed.
pptxDoc.Close();