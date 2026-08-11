using Syncfusion.Presentation;
 
//Opens the source PPTX document. 
IPresentation sourcePptx = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")); 

//Iterates through each section. 
foreach (ISection section in sourcePptx.Sections) 
{
     //Creates a destination PPTX document. Existing presentations can also be used here. 
     IPresentation destinationPptx = Presentation.Create(); 
     // Clone the slides from the section and move to new PPTX document. 
     foreach (ISlide slide in section.Slides) 
     {
         destinationPptx.Slides.Add(slide.Clone(), PasteOptions.SourceFormatting, sourcePptx); 
     }
     //Saves the destination PPTX document. 
     string outputPath = Path.Combine(Path.GetFullPath("Output"), section.Name + "_Slides.pptx");
    destinationPptx.Save(outputPath); 
    destinationPptx.Close(); 
} 
//Closes the PPTX document. 
sourcePptx.Close(); 