using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Copy the column and append it to the end of table.
table.Columns.Add(table.Columns[0].Clone());
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));