using System.Drawing;
using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();

//Adds a slide to the PowerPoint Presentation
ISlide firstSlide = pptxDoc.Slides.Add(SlideLayoutType.Blank);

//Adds a textbox in a slide by specifying its position and size
IShape textShape = firstSlide.AddTextBox(100, 75, 756, 200);

//Adds a paragraph into the textShape
IParagraph paragraph = textShape.TextBody.AddParagraph();

//Set the horizontal alignment of paragraph
paragraph.HorizontalAlignment = HorizontalAlignmentType.Center;

//Adds a textPart in the paragraph
ITextPart textPart = paragraph.AddTextPart("Hello Presentation");

//Applies font formatting to the text
textPart.Font.FontSize = 80;
textPart.Font.Bold = true;

//Adds a new paragraph with text.
paragraph = textShape.TextBody.AddParagraph("AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.");

//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;

//Sets the bullet character for this list
paragraph.ListFormat.BulletCharacter = Convert.ToChar(183);

//Sets the font of the bullet character
paragraph.ListFormat.FontName = "Symbol";

//Sets the hanging value as 20
paragraph.FirstLineIndent = -20;

//Adds a new paragraph  
paragraph = textShape.TextBody.AddParagraph("The company manufactures and sells metal and composite bicycles to North American, European and Asian commercial markets.");

//Sets the list type as bullet
paragraph.ListFormat.Type = ListType.Bulleted;

//Sets the list level as 2. Possible values can range from 0 to 8
paragraph.IndentLevelNumber = 2;

//Gets a picture as stream.
FileStream pictureStream = new FileStream(@"../../../Data/Image.jpg", FileMode.Open);

//Gets the image from file path
Image image = Image.FromStream(pictureStream);

// Adds the image to the slide by specifying position and size
firstSlide.Pictures.AddPicture(new MemoryStream(image.ImageData), 300, 270, 410, 250);

using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
//Releases the resources occupied
pptxDoc.Close();
