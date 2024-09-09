using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Iterate through the rows of the table
foreach (IRow row in table.Rows)
{
	//Iterate through the cells of the rows.
	foreach (ICell cell in row.Cells)
	{
		//Iterate through the paragraphs of the cell.
		foreach (IParagraph paragraph in cell.TextBody.Paragraphs)
		{
			//Iterate through the TextPart of the paragraph.
			foreach (ITextPart textPart in paragraph.TextParts)
			{
				//When the TextPart contains text Panda, then modifies the text content of the TextPart.
				if (textPart.Text.Contains("Panda"))
					textPart.Text = "Hello Presentation";
			}
		}
	}
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);