using Syncfusion.Presentation;

//Creates a new Presentation instance.
IPresentation pptxDoc = Presentation.Create();
//Adds the slide into the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Adds a textbox to hold the list
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 270);
//Opens the image file used as the picture bullet.
FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open, FileAccess.Read);
//Adds a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Sets the list type as Picture
paragraph.ListFormat.Type = ListType.Picture;
//Sets the image for the list.
paragraph.ListFormat.Picture(pictureStream);
//Sets the picture size. Here, 100 means 100% of the text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 150;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Adds another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Sets the list type as Picture
paragraph.ListFormat.Type = ListType.Picture;
//Sets the image for the list.
paragraph.ListFormat.Picture(pictureStream);
//Sets the picture size. Here, 100 means 100% of the text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 150;
//Sets the list level as 1
paragraph.IndentLevelNumber = 1;
//Sets the hanging value
paragraph.FirstLineIndent = -20;
//Saves the Presentation to the file system.
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the picture stream.
pictureStream.Close();
//Closes the Presentation
pptxDoc.Close();