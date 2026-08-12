using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Replace image with text.
ReplaceImageWithText(pptxDoc, "Image was here");
//Save output PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));


void ReplaceImageWithText(IPresentation pptxDoc, string altText)
{
    double shapeLeft, shapeTop, shapeWidth, shapeHeignt = 0;
    //Iterate through the slides collection.
    foreach (ISlide slide in pptxDoc.Slides)
    {
        //Iterate through the pictures collection and remove the picture named "Image".
        foreach (IPicture picture in slide.Pictures)
        {
            if (picture.Description == "ImageID1")
            {
                shapeLeft = picture.Left;
                shapeTop = picture.Top;
                shapeWidth = picture.Width;
                shapeHeignt = picture.Height;
                slide.Pictures.Remove(picture);
                //Insert a textbox with the alternate text in the bounds of picture.
                slide.Shapes.AddTextBox(shapeLeft, shapeTop, shapeWidth, shapeHeignt).TextBody.AddParagraph(altText);
                break;
            }
        }
    }
}