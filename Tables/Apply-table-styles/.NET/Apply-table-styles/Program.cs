using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);

//Add table to the slide.
ITable table = slide.Shapes.AddTable(3, 3, 100, 120, 300, 200);
table.BuiltInStyle = BuiltInTableStyle.ThemedStyle2Accent4;
table.HasBandedRows = false;
table.HasHeaderRow = false;
table.HasBandedColumns = true;
table.HasFirstColumn = true;
table.HasLastColumn = true;
table.HasTotalRow = true;

//Retrieve each cell and fills text content to the cell.
ICell cell = table[0, 0];
cell.TextBody.AddParagraph("First Row and First Column");
cell = table[0, 1];
cell.TextBody.AddParagraph("First Row and Second Column");
cell = table[0, 2];
cell.TextBody.AddParagraph("First Row and Third Column");
cell = table[1, 0];
cell.TextBody.AddParagraph("Second Row and First Column");
cell = table[1, 1];
cell.TextBody.AddParagraph("Second Row and Second Column");
cell = table[1, 2];
cell.TextBody.AddParagraph("Second Row and Third Column");
cell = table[2, 0];
cell.TextBody.AddParagraph("Third Row and First Column");
cell = table[2, 1];
cell.TextBody.AddParagraph("Third Row and Second Column");
cell = table[2, 2];
cell.TextBody.AddParagraph("Third Row and Third Column");
//Add description to table shape.
table.Description = "Table arrangement";
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);