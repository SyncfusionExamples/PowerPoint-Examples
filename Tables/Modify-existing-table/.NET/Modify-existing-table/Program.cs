using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get table from slide.
ITable table = slide.Shapes[0] as ITable;
//Modify the table width.
table.Width = 450;
//Change the built in style of the table.
table.BuiltInStyle = BuiltInTableStyle.DarkStyle1Accent2;
//Set text content to the cell.
table.Rows[0].Cells[0].TextBody.AddParagraph("Row1 Cell1");
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);