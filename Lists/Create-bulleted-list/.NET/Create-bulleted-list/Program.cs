using Syncfusion.Presentation;

//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
// Adds a textbox to hold the list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 250);
//Adds a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as bulleted
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the font for the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the bullet character size. Here, 100 means 100% of the text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;
//Adds another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Sets the list type as bulleted
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the font for the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the bullet character size. Here, 100 means 100% of the text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;
//Adds another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("While its base operation is located in Washington with 290 employees, several regional sales teams are located throughout their market base.");
//Sets the list type as bulleted
paragraph.ListFormat.Type = ListType.Bulleted;
//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the font for the bullet character
paragraph.ListFormat.FontName = "Symbol";
//Sets the bullet character size. Here, 100 means 100% of the text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 100;
//Saves the Presentation to the file system.
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();