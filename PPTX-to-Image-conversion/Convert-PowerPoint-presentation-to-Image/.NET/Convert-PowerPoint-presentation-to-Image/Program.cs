using Syncfusion.Drawing;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Create_PowerPoint_presentation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx"))
            {
                //Initialize the PresentationRenderer to perform image conversion.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert PowerPoint slide to image as stream.
                Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Jpeg);
                //Save the images to file.
                for (int i = 0; i < images.Length; i++)
                {
                    using (Stream stream = images[i])
                    {
                        //Save the image stream to a file.
                        using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("Output/Image" + i + ".jpg")))
                        {
                            stream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}