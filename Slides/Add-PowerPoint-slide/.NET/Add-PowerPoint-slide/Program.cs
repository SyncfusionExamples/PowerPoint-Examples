using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide to the PowerPoint presentation.
ISlide slide = pptxDoc.Slides.Add();
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
