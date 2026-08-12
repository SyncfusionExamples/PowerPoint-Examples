using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a slide of blank layout type.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
