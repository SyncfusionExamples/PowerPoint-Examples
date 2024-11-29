using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Open an existing presentation.
using (FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load the presentation. 
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Initialize PresentationRenderer. 
        pptxDoc.PresentationRenderer = new PresentationRenderer();
        //Convert the PowerPoint slide as an image stream.
        using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
        {
            //Save the image stream to a file.
            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output.jpg")))
            {
                stream.CopyTo(fileStreamOutput);
            }  
        }      
    }       
}