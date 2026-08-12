using Syncfusion.Presentation;

//Opens an existing presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Finds the first occurrence of a particular text in the PowerPoint presentation
    ITextSelection textSelection = pptxDoc.Find("product", false, false);
    //Gets the found text as a single text part
    ITextPart textPart = textSelection.GetAsOneTextPart();
    //Replaces the text
    textPart.Text = "Service";
    //Saves the Presentation
    pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
}