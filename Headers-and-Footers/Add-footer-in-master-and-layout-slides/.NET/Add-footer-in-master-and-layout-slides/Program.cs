using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Access the first master slide in PowerPoint file.
IMasterSlide masterSlide = pptxDoc.Masters[0];
//Set the visibility of Footer content in the Master slide.
masterSlide.HeadersFooters.Footer.Visible = true;
//Set the text to be added to the Footer of the Master slide.
masterSlide.HeadersFooters.Footer.Text = "Master Slide Footer";
//Set the visibility of DateTime Footer in the Master slide.
masterSlide.HeadersFooters.DateAndTime.Visible = true;
//Set the format of the DateTime Footer in the Master slide.
masterSlide.HeadersFooters.DateAndTime.Format = DateTimeFormatType.DateTimehmmssAMPM;
//Iterate each layout slide in the Master slide.
foreach (ILayoutSlide layoutSlide in masterSlide.LayoutSlides)
{
    //Set the visibility of Footer content in the Layout slide.
    layoutSlide.HeadersFooters.Footer.Visible = true;
    //Set the text to be added to the Footer of the Layout slide.
    layoutSlide.HeadersFooters.Footer.Text = "Layout slide Footer";
    //Set the visibility of DateTime Footer in Layout slide.
    layoutSlide.HeadersFooters.DateAndTime.Visible = true;
    //Set the format of the DateTime Footer in Layout slide.
    layoutSlide.HeadersFooters.DateAndTime.Format = DateTimeFormatType.DateTimeMMddyyhmmAMPM;
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);