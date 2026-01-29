using Syncfusion.Drawing;
using Syncfusion.Presentation;

public static class Program
{
    public static int imageIndex = 0;
    public static string outputPath = @"Output";
	
    public static void Main()
    {
        // Open the PowerPoint file as a stream.
        using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
        {
            // Load the presentation from the stream.
            using (IPresentation presentation = Presentation.Open(inputStream))
            {
                // Extract EMF images from Masters and their Layout slides
                foreach (IMasterSlide master in presentation.Masters)
                {
                    // Process shapes placed on the master slide
                    foreach (IShape shape in master.Shapes)
                    {
                        ProcessShape(shape);
                    }
                    // Process shapes placed on each layout slide under this master
                    foreach (ILayoutSlide layoutSlide in master.LayoutSlides)
                    {
                        foreach (IShape shape in layoutSlide.Shapes)
                        {
                            ProcessShape(shape);
                        }
                    }
                }
                // Extract EMF images from Normal slides
                foreach (ISlide slide in presentation.Slides)
                {
                    foreach (IShape shape in slide.Shapes)
                    {
                        ProcessShape(shape);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Processes a shape: saves EMF images and recurses through group shapes.
    /// </summary>
    private static void ProcessShape(IShape shape)
    {
        // If the shape is a picture with EMF format
        if (shape is IPicture picture && picture.ImageFormat == ImageFormat.Emf)
        {
            SaveEMFImage(picture);
        }
        // Shape is not a picture object, but its FILL contains a picture(Picture Fill)
        if (shape.Fill != null && shape.Fill.FillType == FillType.Picture)
        {
            // Validate bytes exist and check if those bytes represent an EMF file
            var bytes = shape.Fill.PictureFill.ImageBytes;
            if (bytes != null && bytes.Length > 0 && IsEmf(bytes))
            {
                string filePath = Path.Combine(outputPath, $"Slide_Image_{++imageIndex}.emf");
                File.WriteAllBytes(filePath, bytes);
            }
        }

        // If the shape is a group, process child shapes
        if (shape is IGroupShape group)
        {
            foreach (IShape child in group.Shapes)
            {
                ProcessShape(child);
            }
        }
    }

    /// <summary>
    /// Saves EMF image from the picture shape.
    /// </summary>
    private static void SaveEMFImage(IPicture picture)
    {
        byte[] imageData = picture.ImageData;

        if (imageData != null && imageData.Length > 0)
        {
            string extension = picture.ImageFormat.ToString().ToLower();
            string filePath = Path.Combine(outputPath, $"Slide_Image_{++imageIndex}.{extension}");
            File.WriteAllBytes(filePath, imageData);
        }
    }
	
    /// <summary>
    /// Checks whether the given byte[] looks like an EMF file.
    /// </summary>
    private static bool IsEmf(byte[] bytes)
    {
        if (bytes == null || bytes.Length < 44) return false;

        // EMF signature " EMF" = 0x20 0x45 0x4D 0x46 (little endian uint32 => 0x464D4520)
        // At offset 40..43
        uint signature = BitConverter.ToUInt32(bytes, 40);
        return signature == 0x464D4520;
    }
}