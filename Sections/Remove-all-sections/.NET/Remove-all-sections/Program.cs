using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Remove the sections.
pptxDoc.Sections.Clear();
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
pptxDoc.Close();