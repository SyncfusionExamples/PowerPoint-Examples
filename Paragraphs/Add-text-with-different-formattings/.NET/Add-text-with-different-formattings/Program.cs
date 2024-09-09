using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Get the slide from Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add textbox to the slide.
IShape textboxShape2 = slide.AddTextBox(500, 0, 400, 500);
//Add paragraph to the textbody of textbox.
IParagraph paragraph2 = textboxShape2.TextBody.AddParagraph();
//Add a TextPart to the paragraph.
ITextPart textPartFormatting = paragraph2.AddTextPart();
//Add text to the TextPart.
textPartFormatting.Text = "In 2000, AdventureWorks Cycles bought a small manufacturing plant, located in Mexico. The plant manufactures several critical subcomponents for the AdventureWorks Cycles product line. These subcomponents are shipped to the another location for final product assembly. In 2001, the plant, became the sole manufacturer and distributor of the touring bicycle product group.";
//Set the underline color.
textPartFormatting.UnderlineColor.SystemColor = Color.Black;
//Retrieve the existing font for modification.
IFont font = textPartFormatting.Font;
//Set the underline type.
font.Underline = TextUnderlineType.Single;
//Set the font weight.
font.Bold = true;
//Add a TextPart to the paragraph.
ITextPart textPartFormatting2 = paragraph2.AddTextPart();
//Add text to the TextPart.
textPartFormatting2.Text = "In 2000, AdventureWorks Cycles bought a small manufacturing plant, located in Mexico.";
//Retrieve the existing font for modification.
IFont font2 = textPartFormatting2.Font;
//Set the font color.
font2.Color.SystemColor = Color.Blue;
//Set the underline type.
font2.Underline = TextUnderlineType.WavyDouble;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);