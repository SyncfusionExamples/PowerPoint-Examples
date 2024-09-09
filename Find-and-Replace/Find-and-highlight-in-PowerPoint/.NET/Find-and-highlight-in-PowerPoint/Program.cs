using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Find all the occurrences of a particular text in the PowerPoint presentation.
ITextSelection[] textSelections = pptxDoc.FindAll("product", false, false);
foreach (ITextSelection textSelection in textSelections)
{
	//Get the found text containing text parts.
	foreach (ITextPart textPart in textSelection.GetTextParts())
	{
		//Set highlight color.
		textPart.Font.HighlightColor = ColorObject.Yellow;
	}
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);