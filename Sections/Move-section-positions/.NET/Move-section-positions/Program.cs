using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Move the second section to third position within the PowerPoint presentation.
pptxDoc.Sections[1].Move(3);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));