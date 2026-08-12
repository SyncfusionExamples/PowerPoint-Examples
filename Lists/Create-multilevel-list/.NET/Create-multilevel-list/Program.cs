using Syncfusion.Presentation;

//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds a textbox to hold the bulleted list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 250);
//Adds paragraph to the textbox
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
//Sets the numbered style (list numbering) as Arabic number following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Sets the starting value as 1
paragraph.ListFormat.StartValue = 1;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Adds paragraph to the textbox
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
//Sets the numbered style (list numbering) as lower case alphabet following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.AlphaLcPeriod;
//Sets the list level as 2
paragraph.IndentLevelNumber = 2;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Adds paragraph to the textbox
paragraph = textBoxShape.TextBody.AddParagraph("While its base operation is located in Washington with 290 employees, several regional sales teams are located throughout their market base.");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
//Sets the numbered style (list numbering) as roman number lower casing following by period.
paragraph.ListFormat.NumberStyle = NumberedListStyle.RomanLcPeriod;
//Sets the list level as 3
paragraph.IndentLevelNumber = 3;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Adds paragraph to the textbox
paragraph = textBoxShape.TextBody.AddParagraph("These subcomponents are shipped to another location for final product assembly");
//Sets the list type as Numbered list
paragraph.ListFormat.Type = ListType.Numbered;
paragraph.ListFormat.NumberStyle = NumberedListStyle.ArabicPeriod;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Saves the Presentation to the file system.
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();
