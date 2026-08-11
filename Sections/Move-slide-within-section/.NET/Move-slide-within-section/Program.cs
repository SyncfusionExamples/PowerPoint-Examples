using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide of second section in the PowerPoint presentation.
ISlide slide = pptxDoc.Sections[1].Slides[0];
//Move the slide to first section.
slide.MoveToSection(0);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));