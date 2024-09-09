using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide with blank layout to presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get the excel file as stream.
using FileStream excelStream = new(Path.GetFullPath("Data/OleTemplate.xlsx"), FileMode.Open);
//Image to be displayed, This can be any image.
using FileStream imageStream = new(Path.GetFullPath("Data/OlePicture.png"), FileMode.Open);
//Add an OLE object to the slide.
IOleObject oleObject = slide.Shapes.AddOleObject(imageStream, "Excel.Sheet.12", excelStream);
//Set size and position of the OLE object.
oleObject.Left = 10;
oleObject.Top = 10;
oleObject.Width = 400;
oleObject.Height = 300;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);