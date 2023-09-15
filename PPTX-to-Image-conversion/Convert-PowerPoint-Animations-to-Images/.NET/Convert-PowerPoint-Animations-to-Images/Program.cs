using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the PresentationAnimationConverter to perform slide to image conversion based on animation order.
using PresentationAnimationConverter animationConverter = new();
int i = 0;
foreach (ISlide slide in pptxDoc.Slides)
{
    //Convert the PowerPoint slide to a series of images based on entrance animation effects.
    Stream[] imageStreams = animationConverter.Convert(slide, ExportImageFormat.Png);

    //Save the image stream.
    foreach (Stream stream in imageStreams)
    {
        i++;
        //Reset the stream position.
        stream.Position = 0;

        //Create the output image file stream.
        using FileStream fileStreamOutput = File.Create("../../../Output" + i + ".png");
        //Copy the converted image stream into created output stream.
        stream.CopyTo(fileStreamOutput);
        //Disposes the stream.
        stream.Dispose();
    }
}