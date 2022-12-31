using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Insert a row at the specified index. Here, the existing first row at index 0 is copied and inserted at row index 1.
table.Rows.Insert(1, table.Rows[0].Clone());
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);