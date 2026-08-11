using Syncfusion.Presentation;

//Open an existing presentation.
using IPresentation presentation = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Encrypt the presentation with a password.
presentation.Encrypt("syncfusion");
presentation.Save(Path.GetFullPath(@"Output/Result.pptx"));