using Syncfusion.Presentation;
using System.Text.RegularExpressions;



//Load or open a PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Input.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);

//Find all the occurrences of a given pattern of text in the PowerPoint presentation using Regex.
ITextSelection[] textSelections = pptxDoc.FindAll(new Regex("{[A-Za-z]+}"));
foreach (ITextSelection textSelection in textSelections)
{
    //Get the found text as a single text part.
    ITextPart textPart = textSelection.GetAsOneTextPart();
    //Replace the text.
    textPart.Text = "Service";
}
//Saves the Presentation.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);


