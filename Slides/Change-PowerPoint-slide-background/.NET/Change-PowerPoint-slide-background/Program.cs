using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using IPresentation pptxDoc = Presentation.Open(inputStream);
    //Retrieves the slide instance.
    ISlide slide = pptxDoc.Slides[0];
    //Retrieves the background instance.
    IBackground background = slide.Background;
    //Sets the fill type of the background to gradient.
    background.Fill.FillType = FillType.Gradient;
    //Retrieves the fill of the background to the IGradientFill instance.
    IGradientFill gradient = background.Fill.GradientFill;
    //Adds the first gradient stop of the gradient fill.
    gradient.GradientStops.Add(ColorObject.Green, 20);
    //Adds the second gradient stop of the gradient fill.
    gradient.GradientStops.Add(ColorObject.Yellow, 50);
    using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    pptxDoc.Save(outputStream);
}