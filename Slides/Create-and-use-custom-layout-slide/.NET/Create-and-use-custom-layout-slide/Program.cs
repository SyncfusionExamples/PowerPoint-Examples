using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Add a new custom layout slide to the master collection with a specific layout type and name.
ILayoutSlide layoutSlide = pptxDoc.Masters[0].LayoutSlides.Add(SlideLayoutType.Blank, "CustomLayout");
//Set background of the layout slide.
layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(78, 89, 90);
//Get a picture as stream.
using (FileStream pictureStream = new(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
//Add the picture into layout slide.
layoutSlide.Shapes.AddPicture(pictureStream, 100, 100, 100, 100);
//Add a slide of new designed custom layout to the presentation.
ISlide slide = pptxDoc.Slides.Add(layoutSlide);
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));
