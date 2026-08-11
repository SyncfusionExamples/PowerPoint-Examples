using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")))
{
    //Initialize PresentationRenderer.
    pptxDoc.PresentationRenderer = new PresentationRenderer();
    //Convert the PowerPoint slide as an image stream.
    using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
    {
        //Save the image stream to a file.
        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Image.jpg")))
        {
            stream.CopyTo(fileStreamOutput);
        }
    }
}