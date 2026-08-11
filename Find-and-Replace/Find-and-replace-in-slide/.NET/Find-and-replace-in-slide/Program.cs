using Syncfusion.Presentation;

//Opens an existing presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    ISlide slide = pptxDoc.Slides[0];
    //Finds all the occurrences of a particular text in the specific slide
    ITextSelection[] textSelections = slide.FindAll("product", false, false);
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