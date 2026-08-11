using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the notes slide from the presenatation. 
INotesSlide notesSlide = pptxDoc.Slides[0].NotesSlide;
//Modify the existing content of the header. 
notesSlide.HeadersFooters.Header.Text = "Header content is modified";
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));