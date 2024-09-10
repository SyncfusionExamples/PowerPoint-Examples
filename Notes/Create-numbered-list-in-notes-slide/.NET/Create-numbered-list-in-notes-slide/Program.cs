using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add new slide with blank slide layout type.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add new notes slide in the specified slide.
INotesSlide notesSlide = slide.AddNotesSlide();
//Add a new paragraph with the text in the left hand side textbox. 
IParagraph paragraph = notesSlide.NotesTextBody.AddParagraph("The Northwind sample database (Northwind.mdb) is included with all versions of Access.");
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
paragraph = notesSlide.NotesTextBody.AddParagraph("It provides data you can experiment with and database objects that demonstrate features you might want to implement in your own databases.");
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
paragraph = notesSlide.NotesTextBody.AddParagraph("Using Northwind, you can become familiar with how a relational database is structured and how the database objects work together to help you enter, store, manipulate, and print your data.");
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
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);