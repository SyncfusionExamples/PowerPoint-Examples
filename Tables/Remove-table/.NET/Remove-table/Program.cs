 using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the table from slide.
ITable table = slide.Shapes[0] as ITable;
//Remove table from shape collection.
slide.Shapes.Remove(table);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));