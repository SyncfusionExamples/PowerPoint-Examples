using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add table to the slide.
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);

//Retrieve first cell.
ICell cell = table[0, 0];
//Retrieve each cell and fills text content to the cell.
cell.TextBody.AddParagraph("First Row and First Column");
cell = table[0, 1];
cell.TextBody.AddParagraph("First Row and Second Column");
cell = table[1, 0];
cell.TextBody.AddParagraph("Second Row and First Column");
cell = table[1, 1];
cell.TextBody.AddParagraph("Second Row and Second Column");

//Retrieve first cell.
cell = table[0, 0];
//Set the column span value to merge the cell.
cell.ColumnSpan = 2;

//Give simple description to table shape.
table.Description = "Table arrangement";
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);