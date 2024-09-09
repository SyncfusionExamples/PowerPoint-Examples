using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the PowerPoint.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add textbox to the slide.
IShape textboxShape = slide.AddTextBox(0, 0, 500, 500);
//Add paragraph to the textbody of textbox.
IParagraph paragraph = textboxShape.TextBody.AddParagraph();
//Add a TextPart to the paragraph.
ITextPart textPart = paragraph.AddTextPart();
//Add text to the TextPart.
textPart.Text = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company. The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets. While its base operation is located in Washington with 290 employees, several regional sales teams are located throughout their market base.";
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);