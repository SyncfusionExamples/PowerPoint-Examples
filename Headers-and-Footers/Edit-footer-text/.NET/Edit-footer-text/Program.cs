using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Gets the first slide from the cloned PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Modify the Footer text.
slide.HeadersFooters.Footer.Text = "Footer content modified";
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));