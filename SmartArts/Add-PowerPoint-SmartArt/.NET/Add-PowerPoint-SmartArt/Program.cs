using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a BasicBlockList SmartArt to the slide at the specified size and position.
ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.BasicBlockList, 0, 0, 640, 426);
//Save the PowerPoint Presentation as stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite))
pptxDoc.Save(outputFileStream);