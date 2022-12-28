using Syncfusion.Presentation;

//Create an instance for PowerPoint presentation
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Mark the presentation as final
pptxDoc.Final = true;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);