using Syncfusion.Presentation;

//Loads or open an PowerPoint Presentation
using (FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using IPresentation pptxDoc = Presentation.Open(inputStream);
    //Accesses the built-in document properties
    Console.WriteLine("Title - {0}", pptxDoc.BuiltInDocumentProperties.Title);
    Console.WriteLine("Author - {0}", pptxDoc.BuiltInDocumentProperties.Author);
    using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    pptxDoc.Save(outputStream);
}
