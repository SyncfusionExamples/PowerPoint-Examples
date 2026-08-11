using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the Slide from Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the SmartArt from Slide.
ISmartArt smartArt = slide.Shapes[0] as ISmartArt;
//Get the first node.
ISmartArtNode firstNode = smartArt.Nodes[0];
// Set the text content of node.
firstNode.TextBody.Paragraphs[0].Text = "Text Modified";
// Set the font color of paragraph.
firstNode.TextBody.Paragraphs[0].Font.Color = ColorObject.Black;
//Set the fill type of node.
firstNode.Shapes[0].Fill.FillType = FillType.Solid;
//Set the fill color of node.
firstNode.Shapes[0].Fill.SolidFill.Color = ColorObject.Lavender;
//Set  transparency value of fill.
firstNode.Shapes[0].Fill.SolidFill.Transparency = 30;
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));