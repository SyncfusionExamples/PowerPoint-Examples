using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Create a new section in the PowerPoint presentation.
pptxDoc.Sections.Add();
//Move the first slide to the created section.
pptxDoc.Slides[0].MoveToSection(0);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));