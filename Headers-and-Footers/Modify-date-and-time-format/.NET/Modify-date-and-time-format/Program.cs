using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Get the first slide from the cloned PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Modify Date and Time format of the Footer.
slide.HeadersFooters.DateAndTime.Format = DateTimeFormatType.DateTimeddddMMMMddyyyy;
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);