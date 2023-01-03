using Syncfusion.Drawing;
using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get the excel file as stream.
using FileStream excelStream = new(@"../../../Data/Book1.xlsx", FileMode.Open);
//Add a chart to the slide with a data range from excel worksheet – excel workbook, worksheet number, Data range, position, and size.
IPresentationChart chart = slide.Charts.AddChart(excelStream, 1, "A1:D4", new RectangleF(100, 10, 700, 500));
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);