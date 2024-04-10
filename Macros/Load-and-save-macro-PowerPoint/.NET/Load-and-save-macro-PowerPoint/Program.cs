using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptm"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Add a blank slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a text box to the slide.
IParagraph paragraph = slide.Shapes.AddTextBox(100, 100, 300, 80).TextBody.AddParagraph("Preserve Macros");
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptm"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);