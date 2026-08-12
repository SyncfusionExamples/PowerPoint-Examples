using Syncfusion.Presentation;

//Open an existing Presentation from file system and it can be decrypted by using the provided password.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"), "MYPASSWORD");
//Get whether the presentation is write Protected. Read - only.
bool writeProtected = pptxDoc.IsWriteProtected;
//Check whether the presentation is write protected.
if (writeProtected)
{
    //Remove the write protection for presentation instance.
    pptxDoc.RemoveWriteProtection();
}
//Save the PowerPoint Presentation file
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));