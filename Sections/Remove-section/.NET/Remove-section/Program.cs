using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Remove the second section from the PowerPoint presentation.
pptxDoc.Sections.Remove(pptxDoc.Sections[1]);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
pptxDoc.Close();