using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
bool matchCase = true;
bool wholeWord = false;
//Find all the occurrences of a particular text which matches the given casing.
ITextSelection[] textSelections = pptxDoc.FindAll("product", matchCase, wholeWord);
foreach (ITextSelection textSelection in textSelections)
{
	//Get the found text as a single text part.
	ITextPart textPart = textSelection.GetAsOneTextPart();
	//Replace the text.
	textPart.Text = "Service";
}
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);