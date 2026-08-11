using Syncfusion.Presentation;

//Open an existing Presentation from file system and it can be decrypted by using the provided password.
using IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"), "syncfusion");
//Decrypt the document.
presentation.RemoveEncryption();
presentation.Save(Path.GetFullPath(@"Output/Result.pptx"));