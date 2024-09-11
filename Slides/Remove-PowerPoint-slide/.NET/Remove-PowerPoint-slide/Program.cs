using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Remove the specified slide from the Presentation.
pptxDoc.Slides.Remove(slide);
// Remove the slide from the specified index.
pptxDoc.Slides.RemoveAt(1);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);