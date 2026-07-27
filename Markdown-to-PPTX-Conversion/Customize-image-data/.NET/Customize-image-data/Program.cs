using Syncfusion.Presentation;
using System.IO;
using System.Net;

namespace Customize_image_data
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create load options.
            LoadOptions loadOptions = new LoadOptions();
            // Specify the format type as Markdown.
            loadOptions.FormatType = FormatType.Markdown;
            // Initialize Markdown import settings for the LoadOptions instance.
            loadOptions.MdImportSettings = new Syncfusion.Office.Markdown.MdImportSettings();
            // Hook the event to customize the image while importing Markdown document.
            loadOptions.MdImportSettings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;
            // Open the Markdown file with load options.
            using (IPresentation presentation = Presentation.Open(Path.GetFullPath("../../../Data/Input.md"), loadOptions))
            {
                // Save as a PowerPoint document.
                presentation.Save(Path.GetFullPath(@"../../../Output/Output.pptx"));
            }
        }
        private static void MdImportSettings_ImageNodeVisited(object sender, Syncfusion.Office.Markdown.MdImageNodeVisitedEventArgs args)
        {
            //Set the image stream based on the image name from the input Markdown.
            if (args.Uri == "Image_1.png")
                args.ImageStream = new FileStream(Path.GetFullPath("../../../Data/Image_1.png"), FileMode.Open);
            else if (args.Uri == "Image_2.png")
                args.ImageStream = new FileStream(Path.GetFullPath("../../../Data/Image_2.png"), FileMode.Open);
            //Retrieve the image from the website and use it.
            else if (args.Uri.StartsWith("https://"))
            {
                WebClient client = new WebClient();
                //Download the image as a stream.
                byte[] image = client.DownloadData(args.Uri);
                Stream stream = new MemoryStream(image);
                //Set the retrieved image from the input Markdown.
                args.ImageStream = stream;
            }
        }
    }
}
