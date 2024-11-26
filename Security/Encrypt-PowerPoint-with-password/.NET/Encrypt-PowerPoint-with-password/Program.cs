using Syncfusion.Presentation;

//Open an existing presentation.
using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read);
using IPresentation presentation = Presentation.Open(inputStream);
//Encrypt the presentation with a password.
presentation.Encrypt("syncfusion");
using FileStream outputStream = new (Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
presentation.Save(outputStream);