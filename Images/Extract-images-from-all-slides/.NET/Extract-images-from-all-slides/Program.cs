using Syncfusion.Presentation;
using System.IO;

namespace Extract_images_from_all_slides
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Presentation.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamPath))
                {
                    int i = 0;
                    //Iterates through the slides collection.
                    foreach (ISlide slide in pptxDoc.Slides)
                    {                       
                        //Iterates through the pictures collection and extract the picture.
                        foreach (IPicture picture in slide.Pictures)
                        {
                            byte[] imageBytes = picture.ImageData;
                            Stream stream = new MemoryStream(imageBytes);
                            //Resets the stream position.
                            stream.Position = 0;
                            string imageFileName = "Image_" + i + "." + picture.ImageFormat.ToString().ToLower();
                            //Creates the output image file stream.
                            using (FileStream fileStreamOutput = new FileStream(Path.GetFullPath(@"../../../Output/" + imageFileName), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Copies the converted image stream into created output stream.
                                stream.CopyTo(fileStreamOutput);
                            }
                            i++;
                        }
                    }
                }
            }
        }
    }
}
