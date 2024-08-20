using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add new slide with blank slide layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add new notes slide in the specified slide.
INotesSlide notesSlide = slide.AddNotesSlide();
//Set the visibility of Header in the Notes slide.
notesSlide.HeadersFooters.Header.Visible = true;
//Set the text to be added to the Header of the Notes slide.
notesSlide.HeadersFooters.Header.Text = "Header is added to Notes slide";
//Set the visibility of DateTime Footer in the Notes slide.
notesSlide.HeadersFooters.DateAndTime.Visible = true;
//Set the format of the DateTime Footer in the Notes slide.
notesSlide.HeadersFooters.DateAndTime.Format = DateTimeFormatType.DateTimeMMMyy;
//Set the visibility of Footer content in the Notes slide.
notesSlide.HeadersFooters.Footer.Visible = true;
//Set the text to be added to the Footer of the Notes slide.
notesSlide.HeadersFooters.Footer.Text = "Notes slide Footer";
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);