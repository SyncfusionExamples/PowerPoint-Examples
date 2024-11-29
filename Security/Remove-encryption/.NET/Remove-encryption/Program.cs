using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing Presentation from file system and it can be decrypted by using the provided password.
using IPresentation presentation = Presentation.Open(inputStream, "syncfusion");
//Decrypt the document.
presentation.RemoveEncryption();
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
presentation.Save(outputStream);