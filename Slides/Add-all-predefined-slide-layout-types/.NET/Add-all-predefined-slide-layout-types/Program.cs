using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();

//Add a slide of Blank type.
ISlide slide1 = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a slide of comparison type.
ISlide slide2 = pptxDoc.Slides.Add(SlideLayoutType.Comparison);
//Add a slide of ContentWithCaption type.
ISlide slide3 = pptxDoc.Slides.Add(SlideLayoutType.ContentWithCaption);
//Add a slide of PictureWithCaption type.
ISlide slide4 = pptxDoc.Slides.Add(SlideLayoutType.PictureWithCaption);
//Add a slide of SectionHeader type.
ISlide slide5 = pptxDoc.Slides.Add(SlideLayoutType.SectionHeader);
//Add a slide of Title type.
ISlide slide6 = pptxDoc.Slides.Add(SlideLayoutType.Title);
//Add a slide of TitleAndContent type.
ISlide slide7 = pptxDoc.Slides.Add(SlideLayoutType.TitleAndContent);
//Add a slide of TitleAndVerticalText type.
ISlide slide8 = pptxDoc.Slides.Add(SlideLayoutType.TitleAndVerticalText);
//Add a slide of TitleOnly type.
ISlide slide9 = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);
//Add a slide of TwoContent type.
ISlide slide10 = pptxDoc.Slides.Add(SlideLayoutType.TwoContent);
//Add a slide of VerticalTitleAndText type.
ISlide slide11 = pptxDoc.Slides.Add(SlideLayoutType.VerticalTitleAndText);

using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
