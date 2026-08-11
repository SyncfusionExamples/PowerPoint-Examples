using Syncfusion.Presentation;

//Open an existing encrypted Presentation file
using IPresentation pptxDoc = Presentation.Open(@"../../../Data/Template.pptx", "password");
//Get the first slide from the PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the first shape of the slide.
IShape shape = slide.Shapes[0] as IShape;
//Change the text of the shape.
if (shape.TextBody.Text == "Company History")
	shape.TextBody.Text = "Company Profile";
pptxDoc.Save(@"../../../Result.pptx");