using Syncfusion.Presentation;

//Opens an existing presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    ISlide slide = pptxDoc.Slides[0];
    //Finds all the occurrences of a particular text in the specific slide
    ITextSelection[] textSelections = slide.FindAll("product", false, false);
    foreach (ITextSelection textSelection in textSelections)
    {
        //Gets the found text containing text parts
        foreach (ITextPart textPart in textSelection.GetTextParts())
        {
            //Sets highlight color
            textPart.Font.HighlightColor = ColorObject.Yellow;
        }
    }
    //Saves the Presentation
    pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
}