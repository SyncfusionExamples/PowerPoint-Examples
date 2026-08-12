using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Creates a Presentation instance
IPresentation pptxDoc = Presentation.Create();
//Adds a blank slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Gets the excel file as stream
FileStream excelStream = new FileStream(Path.GetFullPath("Data/Book1.xlsx"), FileMode.Open);
//Adds a chart to the slide with a data range from excel worksheet – excel workbook, worksheet number, Data range, position, and size.
IPresentationChart chart = slide.Charts.AddChart(excelStream, 1, "A1:D4", new RectangleF(100, 10, 700, 500));
//Saves the Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Closes the Presentation
pptxDoc.Close();