using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
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
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));