using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the layout slide collection of the master.
ILayoutSlides layoutSlides = pptxDoc.Masters[0].LayoutSlides;
ILayoutSlide slideLayout = null;
//Get each layout slide from the collection.
foreach (ILayoutSlide layout in layoutSlides)
{
	//Check if the layout slide has desired custom layout name.
	if (layout.Name == "CustomSlideLayout")
	{
		slideLayout = layout;
		break;
	}
}
//Add slide with the desired layout.
ISlide slide = pptxDoc.Slides.Add(slideLayout);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);