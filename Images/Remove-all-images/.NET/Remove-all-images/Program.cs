using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Iterate through the pictures collection and remove the picture named "Image".
foreach (IPicture picture in slide.Pictures)
{
	//Remove the picture from the slide.
	slide.Pictures.Remove(picture);
	break;
}
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);