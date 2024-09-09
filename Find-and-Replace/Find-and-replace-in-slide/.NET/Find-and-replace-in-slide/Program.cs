using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
ISlide slide = pptxDoc.Slides[0];
//Find all the occurrences of a particular text in the specific slide.
ITextSelection[] textSelections = slide.FindAll("product", false, false);
foreach (ITextSelection textSelection in textSelections)
{
	//Get the found text as a single text part.
	ITextPart textPart = textSelection.GetAsOneTextPart();
	//Replace the text.
	textPart.Text = "Service";
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);