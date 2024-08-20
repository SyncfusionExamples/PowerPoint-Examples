using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Retrieve the first slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Retrieve the first shape from the slide.
IShape shape = slide.Shapes[0] as IShape;
//Get the hyperlink from the shape.
IHyperLink hyperlink = shape.Hyperlink;
//Get the type of action, the hyperlink will be performed when the shape is clicked.
HyperLinkType hyperlinkType = hyperlink.Action;
//Get the target slide of the hyperlink.
ISlide targetSlide = hyperlink.TargetSlide;
//Get the url address of the hyperlink.
string url = hyperlink.Url;
//Get the screen tip text of a hyperlink.
string screenTip = hyperlink.ScreenTip;
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);