using Syncfusion.Presentation;

//Open the existing Presentation file
IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Input.pptx"));
//Hyperlinks to find
string findText1 = "www.xyz.com";
string findText2 = "www.abcd.com";
string findText3 = "www.123.com";
//Get the slide
ISlide slide = pptxDoc.Slides[0];
//Replace the Hyperlink with another hyperlink
ReplaceHyperlink(slide, findText1);
ReplaceHyperlink(slide, findText2);
ReplaceHyperlink(slide, findText3);
//Save the output
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx"));
pptxDoc.Close();

//Replace the hyperlink with another hyperlink
void ReplaceHyperlink(ISlide slide, string findText)
{
    //Replacement Hyperlink
    string replaceText = "www.syncfusion.com";
    //Find the hyperlink from the slide
    ITextSelection[] textSelections = slide.FindAll(findText, true, false);
    if (textSelections != null)
    {
        //Iterate through the textselection
        for (int i = 0; i < textSelections.Length; i++)
        {
            //Get the each textselection as one TextPart
            ITextPart textPart = textSelections[i].GetAsOneTextPart() as ITextPart;
            //Replace the hyperlink text
            textPart.Text = replaceText;
            //Set the new hyperlink
            textPart.SetHyperlink(replaceText);
        }
    }
}