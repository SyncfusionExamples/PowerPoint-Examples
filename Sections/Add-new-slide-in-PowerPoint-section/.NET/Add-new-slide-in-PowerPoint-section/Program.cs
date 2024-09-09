using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a section to the PowerPoint presentation.
ISection section = pptxDoc.Sections.Add();
//Set a name to the created section.
section.Name = "SectionDemo";
//Add a slide to the created section.
ISlide slide = section.AddSlide(SlideLayoutType.Blank);
//Add a text box to the slide.
slide.AddTextBox(10, 10, 150, 100).TextBody.AddParagraph("Slide in SectionDemo");
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);