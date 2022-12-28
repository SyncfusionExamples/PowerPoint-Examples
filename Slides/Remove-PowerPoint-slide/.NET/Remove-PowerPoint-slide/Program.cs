using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using IPresentation pptxDoc = Presentation.Open(inputStream);

    //Retrieves the slide instance.
    ISlide slide = pptxDoc.Slides[0];
    //Removes the specified slide from the Presentation.
    pptxDoc.Slides.Remove(slide);
    // Removes the slide from the specified index.
    pptxDoc.Slides.RemoveAt(1);

    using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    pptxDoc.Save(outputStream);
}
