using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Open a slide to the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the comment from the slide.
IComment comment = slide.Comments[0] as IComment;
//Modify the comment text.
comment.Text = "The comment text content is changed";
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);