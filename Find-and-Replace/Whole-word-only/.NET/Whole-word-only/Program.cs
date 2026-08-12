using Syncfusion.Presentation;

//Opens an existing presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    bool matchCase = false;
    bool wholeWord = true;
    //Finds all the occurrences of a given whole word
    ITextSelection[] textSelections = pptxDoc.FindAll("product", matchCase, wholeWord);
    foreach (ITextSelection textSelection in textSelections)
    {
        //Gets the found text as a single text part
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text
        textPart.Text = "Service";
    }
    //Saves the Presentation
    pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
}