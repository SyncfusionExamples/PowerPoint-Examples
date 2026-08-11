using Syncfusion.Presentation;
using System.Text.RegularExpressions;

//Opens an existing presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Input.pptx")))
{
    //Finds all the occurrences of a given pattern of text in the PowerPoint presentation using Regex.
    ITextSelection[] textSelections = pptxDoc.FindAll(new Regex("{[A-Za-z]+}"));
    foreach (ITextSelection textSelection in textSelections)
    {
        //Gets the found text as a single text part.
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text.
        textPart.Text = "Service";
    }
    //Saves the Presentation.
    pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
}

