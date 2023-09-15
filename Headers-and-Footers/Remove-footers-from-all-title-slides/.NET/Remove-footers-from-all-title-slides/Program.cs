using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
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
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);