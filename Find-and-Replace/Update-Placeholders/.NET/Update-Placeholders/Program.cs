using Syncfusion.Presentation;
using System.ComponentModel;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/UpdatePlaceholders.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

using (IPresentation pptxDoc = Presentation.Open(inputStream))
{
    //Finds all the occurrences of a particular text in the PowerPoint presentation
    ITextSelection[] textSelections1 = pptxDoc.FindAll("{CustomerName}", false, false);
    foreach (ITextSelection textSelection in textSelections1)
    {
        //Gets the found text as a single text part
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text
        textPart.Text = "Nancy Dovolio";
    }
    ITextSelection[] textSelections2 = pptxDoc.FindAll("{Price}", false, false);
    foreach (ITextSelection textSelection in textSelections2)
    {
        //Gets the found text as a single text part
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text
        textPart.Text = "$50";
    }
    // Iterate through all slides in the presentation
    foreach (ISlide slide in pptxDoc.Slides)
    {
        // Iterate through each shape in the slide
        for (int i = 0; i < slide.Shapes.Count; i++)
        {
            IShape shape = slide.Shapes[i] as IShape;

            // Check for picture with specific alt text
            if (shape.SlideItemType is SlideItemType.Placeholder && shape.PlaceholderFormat.Type == PlaceholderType.Picture)
            {
                // Check if the picture has the specific alt text
                if (shape.Description == "{ProfileImage}")
                {
                    // Replace the image
                    FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Customer_profile.png"), FileMode.Open);
                    MemoryStream memoryStream = new MemoryStream();
                    pictureStream.CopyTo(memoryStream);
                    slide.Shapes.AddPicture(memoryStream, shape.Left, shape.Top, shape.Width, shape.Height);
                    slide.Shapes.Remove(shape);
                    // Close the picture stream
                    pictureStream.Close();
                    memoryStream.Close();
                }
            }
        }
    }
    using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    // Save the modified presentation
    pptxDoc.Save(outputStream);
}