using Syncfusion.Office;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add the slide for Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add textbox to the slide.
IShape textboxShape = slide.AddTextBox(500, 0, 400, 500);
//Add paragraph to the textbody of textbox.
IParagraph paragraph = textboxShape.TextBody.AddParagraph();
//Add a TextPart to the paragraph.
ITextPart textPart = paragraph.AddTextPart();
//Add text to the TextPart.
textPart.Text = "Adventure Works Cycles";
//Set a language as "Spanish (Argentina)" for TextPart.
textPart.Font.LanguageID = (short)LocaleIDs.es_AR;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);