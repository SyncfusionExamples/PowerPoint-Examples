using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a SmartArt to the slide at the specified size and position.
ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.AlternatingHexagons, 0, 0, 640, 426);
//Add a new node to the SmartArt.
ISmartArtNode newNode = smartArt.Nodes.Add();
//Set the text to the newly added node.
newNode.TextBody.AddParagraph("New main node added.");
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);