using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Add or append a new row at the end of table.
IRow row = table.Rows.Add();
//Iterate row-wise cells and add text to it.
foreach (ICell cell in row.Cells)
{
	cell.TextBody.AddParagraph(table.Rows.IndexOf(row).ToString());
}
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);