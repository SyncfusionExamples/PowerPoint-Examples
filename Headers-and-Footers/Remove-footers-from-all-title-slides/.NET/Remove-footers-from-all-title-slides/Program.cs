using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath("Data/Template.pptx"));
//Iterate each slide in the Presentation.
foreach (ISlide slide in pptxDoc.Slides)
{
    //Check whether the LayoutType of Layout slide is Title.
    if (slide.LayoutSlide.LayoutType == SlideLayoutType.Title)
    {
        //Set the visibility of DateAndTime in the Title slide.
        slide.HeadersFooters.DateAndTime.Visible = false;
        //Set the visibility of Footer in the Title slide.
        slide.HeadersFooters.Footer.Visible = false;
        //Set the visibility of SlideNumber in the Title slide.
        slide.HeadersFooters.SlideNumber.Visible = false;
    }
}
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));