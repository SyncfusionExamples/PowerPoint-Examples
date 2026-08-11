using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Add or append a new row at the end of table.
IRow row = table.Rows.Add();
//Iterate row-wise cells and add text to it.
foreach (ICell cell in row.Cells)
{
	cell.TextBody.AddParagraph(table.Rows.IndexOf(row).ToString());
}
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));