using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide into the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
// Add a textbox to hold the list.
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 270);

//Add a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Set the list type as Numbered.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Set the starting value as 1.
paragraph.ListFormat.StartValue = 1;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;
//Set the bullet character size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;

//Add another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Set the list type as bulleted.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;
//Set the bullet character size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;

//Add another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("While its base operation is located in Washington with 290 employees, several regional sales teams are located throughout their market base.");
//Set the list type as bulleted.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;
//Set the bullet character size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100; 
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
