using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add new slide with blank slide layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add new notes slide in the specified slide.
INotesSlide notesSlide = slide.AddNotesSlide();
//Add Paragraph into the text body.
IParagraph paragraph = notesSlide.NotesTextBody.AddParagraph();
//Add text part into the Paragraph.
ITextPart textPart = paragraph.AddTextPart();
textPart.Text = "The notes slide represents the contents and key notes of the corresponding slide. It is more useful when we use Presenter View while presenting the seminars through SlideShow.";
//Set Bold format for text content.
textPart.Font.Bold = true;
//Set font style using font name.
textPart.Font.FontName = "Times New Roman";
//Set text content size using FontSize property.
textPart.Font.FontSize = 20;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);