using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
IPresentation pptxDoc = Presentation.Open(inputStream);
//Clone the slides in 3rd section.
ISlides slides = pptxDoc.Sections[2].Clone();
//Create a destination PowerPoint presentation instance. Existing presentations can also be used here.
pptxDoc = Presentation.Create();
//Iterate the cloned slides and adds the slides to the destination presentation.
foreach (ISlide slide in slides)
    pptxDoc.Slides.Add(slide);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
pptxDoc.Close();