using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add table to the slide.
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
//Retrieve each cell and fills text content to the cell.

ICell cell = table[0, 0];
//Set the column width for a cell; this sets the width for entire column.
cell.ColumnWidth = 400;
//Set the margin for the cell.
cell.TextBody.MarginBottom = 0;
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.Orange;
cell.TextBody.AddParagraph("First Row and First Column");

cell = table[0, 1];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.BlueViolet;
cell.TextBody.AddParagraph("First Row and Second Column");

cell = table[1, 0];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.SandyBrown;
cell.TextBody.AddParagraph("Second Row and First Column");

cell = table[1, 1];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.Silver;
cell.TextBody.AddParagraph("Second Row and Second Column");
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);