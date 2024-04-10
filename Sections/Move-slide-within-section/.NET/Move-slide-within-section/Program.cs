using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide of second section in the PowerPoint presentation.
ISlide slide = pptxDoc.Sections[1].Slides[0];
//Move the slide to first section.
slide.MoveToSection(0);
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);