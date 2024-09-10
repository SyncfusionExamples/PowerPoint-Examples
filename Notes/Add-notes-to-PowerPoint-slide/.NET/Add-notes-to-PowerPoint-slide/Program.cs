using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add new slide with blank slide layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add new notes slide in the specified slide.
INotesSlide notesSlide = slide.AddNotesSlide();
//Add text content into the Notes Slide.
notesSlide.NotesTextBody.AddParagraph("Notes content");
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);