using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Create a new section to Insert.
ISection section = pptxDoc.Sections.Add();
//Name the created section.
section.Name = "InsertedSection";
//Insert the section at second position.
pptxDoc.Sections.Insert(1, section);
//Remove the unwanted created section.
pptxDoc.Sections.RemoveAt(pptxDoc.Sections.Count - 1);
//Save the PowerPoint Presentation as stream
using FileStream outputStream = new("../../../Section.pptx", FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
pptxDoc.Close();
