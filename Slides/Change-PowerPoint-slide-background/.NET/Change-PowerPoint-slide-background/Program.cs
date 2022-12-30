using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the slide instance.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the background instance.
IBackground background = slide.Background;
//Set the fill type of the background to gradient.
background.Fill.FillType = FillType.Gradient;
//Retrieve the fill of the background to the IGradientFill instance.
IGradientFill gradient = background.Fill.GradientFill;
//Add the first gradient stop of the gradient fill.
gradient.GradientStops.Add(ColorObject.Green, 20);
//Add the second gradient stop of the gradient fill.
gradient.GradientStops.Add(ColorObject.Yellow, 50);
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);