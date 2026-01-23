using Syncfusion.Presentation;
using Syncfusion.Drawing;

public static class Program
{
    public static void Main()
    {
        // Open the PowerPoint file as a stream.
        using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx", FileMode.Open, FileAccess.Read))
        {
            // Load the presentation from the stream.
            using (IPresentation presentation = Presentation.Open(inputStream))
            {
                int imageIndex = 0;

                // Iterate through each slide in the presentation.
                foreach (ISlide slide in presentation.Slides)
                {
                    // Iterate through each shape in the slide.
                    foreach (IShape shape in slide.Shapes)
                    {
                        // Check if the shape is an image.
                        if (shape is IPicture picture && picture.ImageFormat== ImageFormat.Emf)
                        {
                            // Retrieve the raw image data.
                            byte[] imageData = picture.ImageData;

                            // Proceed only if image data is valid.
                            if (imageData != null && imageData.Length > 0)
                            {
                                // Determine the image format extension.
                                string extension = picture.ImageFormat.ToString();
                                string outputPath = Path.Combine($"Output/Slide{slide.SlideNumber}_Image{++imageIndex}.{extension}");
                                // Save the image data to the specified path.
                                File.WriteAllBytes(outputPath, imageData);
                            }
                        }
                    }
                }
            }
        }
    }
}
