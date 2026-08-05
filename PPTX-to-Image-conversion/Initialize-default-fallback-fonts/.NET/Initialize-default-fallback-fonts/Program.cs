using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open the existing PowerPoint presentation.
using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
{
    //Initialize the PresentationRenderer to perform image conversion.
    pptxDoc.PresentationRenderer = new PresentationRenderer();
    //Use a set of default FallbackFont collections for IPresentation.
    pptxDoc.FontSettings.FallbackFonts.InitializeDefault();
    //Convert PowerPoint slide to image as stream.
    using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
    {
        //Reset the stream position.
        stream.Position = 0;
        //Create the output image file stream.
        using (FileStream fileStreamOutput = File.Create("Output/Output.jpg"))
        {
            //Copy the converted image stream into created output stream.
            stream.CopyTo(fileStreamOutput);
        }
    }
}