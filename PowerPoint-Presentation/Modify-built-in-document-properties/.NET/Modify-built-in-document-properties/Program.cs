using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using IPresentation pptxDoc = Presentation.Open(inputStream);
    //Modifies the Built-in document properties
    pptxDoc.BuiltInDocumentProperties.Category = "Sales reports";
    pptxDoc.BuiltInDocumentProperties.Company = "Northwind traders";
    using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    pptxDoc.Save(outputStream);
}
