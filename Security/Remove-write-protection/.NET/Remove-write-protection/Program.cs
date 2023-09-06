using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing Presentation from file system and it can be decrypted by using the provided password.
using IPresentation pptxDoc = Presentation.Open(inputStream, "MYPASSWORD");
//Get whether the presentation is write Protected. Read - only.
bool writeProtected = pptxDoc.IsWriteProtected;
//Check whether the presentation is write protected.
if (writeProtected)
{
    //Remove the write protection for presentation instance.
    pptxDoc.RemoveWriteProtection();
}
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);