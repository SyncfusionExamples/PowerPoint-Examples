using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptm"));
//Check whether the presentation has macros and then removes them.
if (pptxDoc.HasMacros)
    pptxDoc.RemoveMacros();
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));