using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a table to the slide
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
//Initialize index values to add text to table cells
int row , col = 0;
//Iterate row-wise cells and add text to it
foreach (IColumn columns in table.Columns)
{
    row = 0;
    foreach (ICell cell in columns.Cells)
    {
        cell.TextBody.AddParagraph("(" + row.ToString() + " , " + col.ToString() + ")");
        row++;
    }
    col++;
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);