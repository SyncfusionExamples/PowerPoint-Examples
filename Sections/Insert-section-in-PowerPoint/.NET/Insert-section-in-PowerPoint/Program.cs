using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Create a new section to Insert.
ISection section = pptxDoc.Sections.Add();
//Name the created section.
section.Name = "InsertedSection";
//Insert the section at second position.
pptxDoc.Sections.Insert(1, section);
//Remove the unwanted created section.
pptxDoc.Sections.RemoveAt(pptxDoc.Sections.Count - 1);
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Section.pptx"));
pptxDoc.Close();
