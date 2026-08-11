using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Accesse an existing custom document property.
IDocumentProperty property = pptxDoc.CustomDocumentProperties["PropertyA"];
//Modify the value of DocumentProperty.
property.Value = "Hello world";
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));