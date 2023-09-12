using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the slide from the Presentation.
ISlide slide = pptxDoc.Slides[0];
//Get the comment in the slide.
IComment comment = slide.Comments[0] as IComment;
//Add reply to the comment.
slide.Comments.Add("Author2", "A2", "Yes, we can we change the font size to 20", DateTime.Now, comment);
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);