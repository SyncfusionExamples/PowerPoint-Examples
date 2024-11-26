using Syncfusion.Drawing;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Create_PowerPoint_presentation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint to image as stream.
                    Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Jpeg);
                    //Saves the images to file system
                    for (int i = 0; i < images.Length; i++)
                    {
                        using (Stream stream = images[i])
                        {
                            //Create the output image file stream
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("Output/Output" + i + ".jpg")))
                            {
                                //Copy the converted image stream into created output stream
                                stream.CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
    }
}