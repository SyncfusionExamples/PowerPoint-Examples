using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add the slide into the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a textbox to hold the list.
IShape textBoxShape = slide.AddTextBox(65, 140, 410, 270);

//Get a picture as stream.
using FileStream pictureStream = new(@"../../../Data/Image.jpg", FileMode.Open);
//Add a new paragraph with the text in the left hand side textbox.
IParagraph paragraph = textBoxShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Set the list type as Numbered.
paragraph.ListFormat.Type = ListType.Picture;
//Set the image for the list.
paragraph.ListFormat.Picture(pictureStream);
//Set the picture size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 150;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;

//Add another paragraph with the text in the left hand side textbox.
paragraph = textBoxShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Set the list type as picture.
paragraph.ListFormat.Type = ListType.Picture;
//Set the image for the list.
paragraph.ListFormat.Picture(pictureStream);
//Set the picture size. Here, 100 means 100% of its text. Possible values can range from 25 to 400.
paragraph.ListFormat.Size = 150;
//Set the list level as 1.
paragraph.IndentLevelNumber = 1;
//Set the hanging value.
paragraph.FirstLineIndent = -20;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);