using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Insert a column at the specified index. Here, the existing first column at index 0 is copied and inserted at column index 1.
table.Columns.Insert(1, table.Columns[0].Clone());
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));