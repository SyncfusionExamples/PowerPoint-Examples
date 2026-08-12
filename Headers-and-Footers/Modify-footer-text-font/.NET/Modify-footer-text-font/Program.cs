using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath("Data/Template.pptx"));
//Get the first slide from the cloned PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Iterate each shape in slide.
foreach (IShape shape in slide.Shapes)
{
    //Check whether the shape is with Placeholder SlideItemType and PlaceholderType as Footer
    if (shape.SlideItemType == SlideItemType.Placeholder && shape.PlaceholderFormat.Type == PlaceholderType.Footer)
    {
        //Change the font name for the Footer content.
        shape.TextBody.Paragraphs[0].Font.FontName = "Verdana";
        //Change the font size for the Footer content.
        shape.TextBody.Paragraphs[0].Font.FontSize = 18;
    }
}
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));