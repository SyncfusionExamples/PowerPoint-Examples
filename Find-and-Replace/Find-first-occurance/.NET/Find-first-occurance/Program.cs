using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Find the first occurrence of a particular text in the PowerPoint presentation.
ITextSelection textSelection = pptxDoc.Find("product", false, false);
//Get the found text as a single text part.
ITextPart textPart = textSelection.GetAsOneTextPart();
//Replace the text.
textPart.Text = "Service";
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);