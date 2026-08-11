using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape from the slide.
IShape shape = slide.Shapes[0] as IShape;
//Remove the hyperlink from the shape.
shape.RemoveHyperlink();
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));