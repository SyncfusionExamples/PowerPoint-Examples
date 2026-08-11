using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Clone the slides in 3rd section.
ISlides slides = pptxDoc.Sections[2].Clone();
//Create a destination PowerPoint presentation instance. Existing presentations can also be used here.
pptxDoc = Presentation.Create();
//Iterate the cloned slides and adds the slides to the destination presentation.
foreach (ISlide slide in slides)
    pptxDoc.Slides.Add(slide);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
pptxDoc.Close();