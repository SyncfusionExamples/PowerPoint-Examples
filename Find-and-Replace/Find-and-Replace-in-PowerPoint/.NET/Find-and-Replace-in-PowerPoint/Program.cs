using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Find all the occurrences of a particular text in the PowerPoint presentation.
ITextSelection[] textSelections = pptxDoc.FindAll("product", false, false);
foreach (ITextSelection textSelection in textSelections)
{
	//Get the found text as a single text part.
	ITextPart textPart = textSelection.GetAsOneTextPart();
	//Replace the text.
	textPart.Text = "Service";
}
//Saves the Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));