using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide to the PowerPoint presentation.
ISlide slide = pptxDoc.Slides.Add();
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
