using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using IPresentation pptxDoc = Presentation.Open(inputStream);
    //Retrieves the slide instance.
    ISlide slide = pptxDoc.Slides[0];
    //Creates a cloned copy of slide.
    ISlide slideClone = slide.Clone();
    //Adds a new text box to the cloned slide.
    IShape textboxShape = slideClone.AddTextBox(0, 0, 250, 250);
    //Adds a paragraph with text content to the shape.
    textboxShape.TextBody.AddParagraph("Hello Presentation");
    //Adds the slide to the Presentation.
    pptxDoc.Slides.Add(slideClone);
    using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    pptxDoc.Save(outputStream);
}