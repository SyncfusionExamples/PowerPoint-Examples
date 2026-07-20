using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.Drawing;
using System.Drawing.Imaging;
using Image = System.Drawing.Image;
using ImageFormat = System.Drawing.Imaging.ImageFormat;

namespace ShowWarningsSample
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            using (FileStream fileStreamInput = new FileStream(@"../../../Data/Sample.pptx", FileMode.Open, FileAccess.Read))
            {
                //Open the existing PowerPoint presentation with stream.
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    //Convert EMF images to PNG
                    ConvertEMFToPNG(pptxDoc);
                    //Initialize the PresentationRenderer to perform image conversion.
                    pptxDoc.PresentationRenderer = new PresentationRenderer();

                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        //Convert the PowerPoint document to PDF document.
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                        {
                            //Save the converted PDF document to MemoryStream.
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;
                        }
                        //Create the output PDF file stream
                        using (FileStream fileStreamOutput = File.Create(@"../../../Data/PPTXToPDF.pdf"))
                        {
                            //Copy the converted PDF stream into created output PDF stream
                            pdfStream.CopyTo(fileStreamOutput);
                        }
                    }

                }
            }
        }


        /// <summary>
        /// Iterate through each slide in the Presentation.
        /// </summary>
        /// <param name="pptxDoc">The Presentation document.</param>
        public static void ConvertEMFToPNG(IPresentation pptxDoc)
        {
            //Iterate in master slides.
            foreach (IMasterSlide masterSlide in pptxDoc.Masters)
            {
                foreach (ILayoutSlide layoutSlide in masterSlide.LayoutSlides)
                {
                    //Iterate in layout slide.
                    IterateShapes(layoutSlide.Shapes);
                }
                IterateShapes(masterSlide.Shapes);
            }
            //Iterate in slides.
            foreach (ISlide slide in pptxDoc.Slides)
            {
                //Iterate in layout slide.
                IterateShapes(slide.LayoutSlide.Shapes);
                IterateShapes(slide.Shapes);
            }
        }

        /// <summary>
        /// Iterate through each shapes in the slide.
        /// </summary>
        /// <param name="shapes">The shapes.</param>
        public static void IterateShapes(IShapes shapes)
        {
            //Iterate through each shape in the slide.
            for (int i = 0; i < shapes.Count; i++)
            {
                IShape shape = (IShape)shapes[i];
                //Identify the slide item type of the shape.
                switch (shape.SlideItemType)
                {
                    case SlideItemType.Picture:
                        IPicture picture = shape as IPicture;
                        if (picture.ImageData != null)
                        {
                            Image image = Image.FromStream(new MemoryStream(picture.ImageData));
                            if (image.RawFormat.Equals(ImageFormat.Emf))
                            {
                                // Store original properties
                                double left = picture.Left;
                                double top = picture.Top;
                                double width = picture.Width;
                                double height = picture.Height;
                                // Convert EMF to PNG stream
                                MemoryStream stream;
                                stream = ConvertEmfToPng(picture.ImageData);

                                // Add replacement picture
                                IPicture newPicture = shapes.AddPicture(stream, left, top, width, height);
                                newPicture.Left = left;
                                newPicture.Top = top;
                                // Remove original EMF picture
                                shapes.Remove(picture);
                                image.Dispose();
                            }
                        }
                        break;
                    case SlideItemType.GroupShape:
                        IGroupShape groupShape = (IGroupShape)shape;
                        //Iterate the shape collection in a group shape.
                        IterateShapes(groupShape.Shapes);
                        break;
                    case SlideItemType.AutoShape:
                    case SlideItemType.Placeholder:
                        IShape autoShape = shape as IShape;
                        //Check the fill type of the shape.
                        if (autoShape.Fill.FillType == FillType.Picture)
                        {
                            //Get the instance for picture Fill.
                            IPictureFill pictureFill = autoShape.Fill.PictureFill;
                            Image fillImage = Image.FromStream(new MemoryStream(pictureFill.ImageBytes));
                            if (fillImage.RawFormat.Equals(ImageFormat.Emf))
                            {
                                byte[] imageBytes = shape.Fill.PictureFill.ImageBytes;

                                using Image image = Image.FromStream(new MemoryStream(imageBytes));

                                // Process only EMF images
                                if (image.RawFormat.Guid != ImageFormat.Emf.Guid)
                                    return;

                                // Convert EMF -> PNG
                                using MemoryStream pngStream = ConvertEmfToPng(imageBytes);

                                // Save existing properties
                                double left = shape.Left;
                                double top = shape.Top;
                                double width = shape.Width;
                                double height = shape.Height;

                                // Create replacement shape
                                IShape newShape = shapes.AddShape(shape.AutoShapeType, left, top, width, height);

                                // Apply PNG picture fill
                                newShape.Fill.FillType = FillType.Picture;
                                newShape.Fill.PictureFill.ImageBytes = pngStream.ToArray();

                                // Copy rotation
                                newShape.Rotation = shape.Rotation;

                                // Remove original shape
                                shapes.Remove(shape);

                            }
                        }
                        break;
                    case SlideItemType.OleObject:
                        IOleObject oleObject = shape as IOleObject;

                        Image oleImage = Image.FromStream(new MemoryStream(oleObject.ImageData));
                        if (oleImage.RawFormat.Equals(ImageFormat.Emf))
                        {
                            double height = oleObject.Height;
                            double width = oleObject.Width;
                            double left = oleObject.Left;
                            double top = oleObject.Top;

                            // Convert preview image to PNG stream
                            MemoryStream pngStream = new MemoryStream();
                            oleImage.Save(pngStream, ImageFormat.Png);
                            pngStream.Position = 0;

                            // Get OLE data
                            byte[] objectData = oleObject.ObjectData;
                            MemoryStream objectStream = new MemoryStream(objectData);

                            // Get ProgID
                            string progID = oleObject.ProgID;

                            // Remove the existing OLE object
                            shapes.Remove(oleObject);

                            // Add a new OLE object with PNG preview image
                            IOleObject replacedOleObject = shapes.AddOleObject(pngStream, progID, objectStream);

                            // Restore position and size
                            replacedOleObject.Left = left;
                            replacedOleObject.Top = top;
                            replacedOleObject.Width = width;
                            replacedOleObject.Height = height;

                            // Dispose
                            objectStream.Dispose();
                            pngStream.Dispose();
                            oleImage.Dispose();
                        }

                        break;
                }
            }
        }
        /// <summary>
        /// Convert EMF to PNG
        /// </summary>
        private static MemoryStream ConvertEmfToPng(byte[] emfBytes)
        {
            using MemoryStream emfStream = new MemoryStream(emfBytes);
            using Metafile metafile = new Metafile(emfStream);

            GraphicsUnit pageUnit = GraphicsUnit.Pixel;
            RectangleF bounds = metafile.GetBounds(ref pageUnit);

            int width = (int)Math.Ceiling(bounds.Width);
            int height = (int)Math.Ceiling(bounds.Height);

            MemoryStream pngStream = new MemoryStream();

            using (Bitmap bitmap = new Bitmap(width, height))
            {
                bitmap.SetResolution(300, 300);

                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Transparent);

                    graphics.SmoothingMode =
                                    System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    graphics.InterpolationMode =
                        System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode =
                        System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                    graphics.DrawImage(
                        metafile,
                        new Rectangle(0, 0, width, height));
                }

                bitmap.Save(pngStream, ImageFormat.Png);
            }

            pngStream.Position = 0;
            return pngStream;

        }
    }
}