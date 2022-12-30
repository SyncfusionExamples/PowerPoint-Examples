using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add the slide into the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a textbox to hold the bulleted list.
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 250);

//Add paragraph to the textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Set the list type as Numbered list.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Set the starting value as 1.
paragraph.ListFormat.StartValue = 1;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;

//Add paragraph to the textbox.
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Set the list type as Numbered list.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as lower case alphabet following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.AlphaLcPeriod;
//Set the list level as 2.
paragraph.IndentLevelNumber = 2;
//Set the hanging value.
paragraph.FirstLineIndent = -20;

//Add paragraph to the textbox.
paragraph = textBoxShape.TextBody.AddParagraph("While its base operation is located Washington with 290 employees, several regional sales teams are located throughout their market base.");
//Set the list type as Numbered list.
paragraph.ListFormat.Type = ListType.Numbered;
//Set the numbered style (list numbering) as roman number lower casing following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.RomanLcPeriod;
//Set the list level as 3.
paragraph.IndentLevelNumber = 3;
//Set the hanging value.
paragraph.FirstLineIndent = -20;

//Add paragraph to the textbox.
paragraph = textBoxShape.TextBody.AddParagraph("These subcomponents are shipped to another location for final product assembly");
//Set the list type as Numbered list.
paragraph.ListFormat.Type = ListType.Numbered;
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
