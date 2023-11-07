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
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStream))
                {
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();
                    //Convert PowerPoint slide to image as stream.
                    Stream stream = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Presentation.ExportImageFormat.Jpeg);
                    //Reset the stream position.
                    stream.Position = 0;
                    //Create the output image file stream.
                    using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../PPTXToImage.jpeg")))
                    {
                        //Copy the converted image stream into created output stream.
                        stream.CopyTo(fileStreamOutput);
                    }
                }
            }
            // Open the PowerPoint located at the specified path using the default associated program.
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../PPTXToImage.jpeg"))
            {
                UseShellExecute = true
            };
            process.Start();
        }
    }
}