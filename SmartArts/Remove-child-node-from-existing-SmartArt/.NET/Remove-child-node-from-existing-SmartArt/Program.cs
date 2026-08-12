using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the SmartArt from slide.
ISmartArt smartArt = slide.Shapes[0] as ISmartArt;
//Remove a node at the specified index.
smartArt.Nodes.RemoveAt(4);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));