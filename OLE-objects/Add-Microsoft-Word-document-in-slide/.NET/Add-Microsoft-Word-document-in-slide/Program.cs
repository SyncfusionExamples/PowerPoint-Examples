using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide with blank layout to presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Get the Word document file as stream.
using FileStream wordDocumentStream = new(Path.GetFullPath(@"Data/OleTemplate.docx"), FileMode.Open);
//Image to be displayed, This can be any image.
using FileStream imageStream = new(Path.GetFullPath(@"Data/OlePicture.png"), FileMode.Open);
//Add an OLE object to the slide.
IOleObject oleObject = slide.Shapes.AddOleObject(imageStream, "Word.Document.12", wordDocumentStream);
//Set size and position of the OLE object.
oleObject.Left = 10;
oleObject.Top = 10;
oleObject.Width = 400;
oleObject.Height = 300;
//Set DisplayAsIcon as true, to open the embedded document in separate (default) application.
oleObject.DisplayAsIcon = true;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);