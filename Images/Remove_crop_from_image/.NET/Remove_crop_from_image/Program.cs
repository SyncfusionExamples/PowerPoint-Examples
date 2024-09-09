using Syncfusion.Presentation;
using SkiaSharp;

//Open an existing PowerPoint Presentation.
using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.pptx"), FileMode.Open, FileAccess.Read);
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first picture from the slide.
IPicture picture = slide.Pictures[0];
//Remove the crop from image.
SKBitmap bitmap = SKBitmap.Decode(new MemoryStream(picture.ImageData));
//Reset the picture with original width and height.
picture.Crop.Width = PixelToPoint(bitmap.Width);
picture.Crop.Height = PixelToPoint(bitmap.Height);
picture.Crop.ContainerWidth = PixelToPoint(bitmap.Width);
picture.Crop.ContainerHeight = PixelToPoint(bitmap.Height);
picture.Crop.OffsetX = 0;
picture.Crop.OffsetY = 0;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create);
pptxDoc.Save(outputStream);


/// <summary>
/// Convert a value from pixel to point.
/// </summary>
/// <param name="value">The value in pixel which needs to be converted to point.</param>
/// <returns>The value converted to point.</returns>
static int PixelToPoint(double value)
{
    return Convert.ToInt32((value * 72) / 96);
}