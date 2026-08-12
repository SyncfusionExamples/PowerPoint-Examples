using Syncfusion.Presentation;

//Open an existing PowerPoint Presentation
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Retrieve the first slide from the Presentation
ISlide slide = pptxDoc.Slides[0];
//Iterate through a copy of the pictures collection and remove each picture.
foreach (IPicture picture in slide.Pictures)
{
    //Remove the picture from the slide.
    slide.Pictures.Remove(picture);
}
//Save the PowerPoint Presentation to a file
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
//Close the Presentation
pptxDoc.Close();