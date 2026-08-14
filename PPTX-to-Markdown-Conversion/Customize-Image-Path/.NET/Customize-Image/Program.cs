using Syncfusion.Office.Markdown;
using Syncfusion.Presentation;
using System.IO;

namespace Customize_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("Data/Input.pptx")))
            {
                // Hook the event to customize the image.
                presentation.MdSaveOptions.ImageNodeVisited += SaveImage;
                // Save the PowerPoint Presentation as a Markdown file.
                presentation.Save(Path.GetFullPath(@"Output/Output.md"));
            }
        }
        /// <summary>
        /// Saves the extracted image to a file and sets the image URI
        /// for use in the generated Markdown output.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">
        /// Contains the image stream and properties used during Markdown image processing.
        /// </param>
        static void SaveImage(object sender, MdImageNodeVisitedEventArgs args)
        {
            string imagepath = Path.GetFullPath(@"Output/Image.png");
            //Save the image stream as a file.
            using (FileStream fileStreamOutput = File.Create(imagepath))
                args.ImageStream.CopyTo(fileStreamOutput);
            //Set the URI to be used for the image in the output Markdown. 
            args.Uri = imagepath;
        }
    }
}