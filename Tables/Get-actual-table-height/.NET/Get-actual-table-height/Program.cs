using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Initialize Presentation renderer.
pptxDoc.PresentationRenderer = new PresentationRenderer();
//Get a table in the slide.
ITable table = pptxDoc.Slides[0].Shapes[0] as ITable;
//Changing the paragraph content in the table.
table.Rows[0].Cells[0].TextBody.AddParagraph("Hello World");
//Get the dynamic height of the table.
float height = table.GetActualHeight();
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));