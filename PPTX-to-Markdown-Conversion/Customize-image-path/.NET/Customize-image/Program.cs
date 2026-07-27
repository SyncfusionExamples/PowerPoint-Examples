using Syncfusion.Office.Markdown;
using Syncfusion.Presentation;
using System.IO;

namespace Customize_image
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Presentation document.
            using (IPresentation presentation = Presentation.Open("../../../Data/Input.pptx"))
            {
                // Hook the event to customize the image.
                presentation.MdSaveOptions.ImageNodeVisited += SaveImage;
                // Save the PowerPoint Presentation as a Markdown file.
                presentation.Save(@"../../../Output/Output.md");
            }
        }
        static void SaveImage(object sender, MdImageNodeVisitedEventArgs args)
        {
            string imagepath = @"../../../Output/Image.png";
            //Save the image stream as a file.
            using (FileStream fileStreamOutput = File.Create(imagepath))
                args.ImageStream.CopyTo(fileStreamOutput);
            //Set the URI to be used for the image in the output Markdown. 
            args.Uri = imagepath;
        }
    }
}