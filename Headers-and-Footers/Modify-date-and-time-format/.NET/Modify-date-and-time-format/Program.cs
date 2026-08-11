using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Get the first slide from the cloned PowerPoint presentation.
ISlide slide = pptxDoc.Slides[0];
//Modify Date and Time format of the Footer.
slide.HeadersFooters.DateAndTime.Format = DateTimeFormatType.DateTimeddddMMMMddyyyy;
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));