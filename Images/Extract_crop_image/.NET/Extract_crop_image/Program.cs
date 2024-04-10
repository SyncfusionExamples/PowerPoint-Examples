using SkiaSharp;
using Syncfusion.Office;
using Syncfusion.Presentation;
using System.IO;


//Open an existing PowerPoint Presentation.
using FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Sample.pptx"), FileMode.Open, FileAccess.Read);
using IPresentation presentation = Presentation.Open(fileStream);
//Get the first slide.
ISlide slide = presentation.Slides[0];
//Get the picture.
IPicture picture = slide.Pictures[0];
//Get cropped image stream.
using MemoryStream stream = GetCroppedPictureData(new MemoryStream(picture.ImageData), picture.Crop);
//Save the image as stream.
using FileStream fileStreamOutput = new FileStream(Path.GetFullPath(@"../../../Output.png"), FileMode.Create);
stream.CopyTo(fileStreamOutput);

// Get the cropped picture data stream.
static MemoryStream GetCroppedPictureData(MemoryStream imageStream, ICrop crop)
{
    //Get the crop properties.
    float imageWidth = crop.ContainerWidth;
    float imageHeight = crop.ContainerHeight;
    float offsetX = crop.OffsetX;
    float offsetY = crop.OffsetY;
    float rectWidth = crop.Width;
    float rectHeight = crop.Height;
    float rectX = (imageWidth - rectWidth) / 2;
    float rectY = (imageHeight - rectHeight) / 2;
    //Creating memory stream to return the cropped image stream.
    MemoryStream outputStream = new MemoryStream();
    //Create SKBitmap with respect to given image size.
    using SKBitmap m_sKBitmap = new SKBitmap(PointToPixel(imageWidth), PointToPixel(imageHeight), SKImageInfo.PlatformColorType, SKAlphaType.Premul);
    //Create SKSurface from the SKBitmap.
    using SKSurface m_surface = SKSurface.Create(m_sKBitmap.Info);
    //Draw the given image with in the image canvas.
    using SKCanvas m_canvas = m_surface.Canvas;
    //Set the background color of the image as transparent.
    m_canvas.Clear(new SKColor(0, 0, 0, 0));
    //Create SKBitmap for the given image data.
    using SKBitmap inputSkBitmap = SKBitmap.Decode(SKData.Create(imageStream));
    // Translate the graphics and apply the Crop.OffsetX and Crop.OffsetY values.
    m_canvas.Translate(PointToPixel(offsetX), PointToPixel(offsetY));
    // Draw the cropped image at the specified picture position.
    SKRect srcRect = new SKRect(PointToPixel(rectX), PointToPixel(rectY), PointToPixel(rectX + rectWidth), PointToPixel(rectY + rectHeight));
    // Apply the crop values.
    m_canvas.DrawBitmap(inputSkBitmap, srcRect, new SKPaint() { FilterQuality = SKFilterQuality.High });

    //Save the SKSurface as stream.
    m_surface.Snapshot().Encode(SKEncodedImageFormat.Png, 100).SaveTo(outputStream);
    outputStream.Position = 0;
    return outputStream;
}

/// <summary>
/// Convert a value from point to pixel.
/// </summary>
/// <param name="value">The value in point which needs to be converted to pixel.</param>
/// <returns>The value converted to pixel.</returns>
static int PointToPixel(double value)
{
    return Convert.ToInt32((value * 96) / 72);
}
