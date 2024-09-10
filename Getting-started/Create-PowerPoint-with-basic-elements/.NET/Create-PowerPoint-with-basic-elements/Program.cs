using System.Drawing;
using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();

//Add a slide to the PowerPoint Presentation.
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a textbox in a slide by specifying its position and size.
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);

//Add a paragraph into the textShape.
IParagraph paragraph = textShape.TextBody.AddParagraph();
//Set the horizontal alignment of paragraph.
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;
//Add a textPart in the paragraph.
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");
//Apply font formatting to the text.
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;
//Add a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");
//Set the list type as bullet.
paragraph.ListFormat.Type = ListType.Bulleted;
//Set the bullet character for this list.
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);
//Set the font of the bullet character.
paragraph.ListFormat.FontName = "Symbol";
//Set the hanging value as 20.
paragraph.FirstLineIndent = -20;
//Add a new paragraph.
paragraph = textShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");
//Set the list type as bullet.
paragraph.ListFormat.Type = ListType.Bulleted;
//Set the list level as 2. Possible values can range from 0 to 8.
paragraph.IndentLevelNumber = 2;

//Get a picture as stream.
using FileStream pictureStream = new(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Get the image from file path.
Image image = Image.FromStream(pictureStream);
// Add the image to the slide by specifying position and size.
firstSlide.Pictures.AddPicture(new MemoryStream(image.ImageData), 300, 270, 410, 250);

using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);