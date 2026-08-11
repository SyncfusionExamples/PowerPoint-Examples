using Syncfusion.Presentation;

//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));
//Traverse through shape in the first slide.
foreach (IShape shape in pptxDoc.Slides[0].Shapes)
{
    if (shape is ISmartArt)
    {
        //Traverse through all nodes inside SmartArt.
        foreach (ISmartArtNode mainNode in (shape as ISmartArt).Nodes)
        {
            if (mainNode.TextBody.Text == "Old Content")
                //Change the node content.
                mainNode.TextBody.Paragraphs[0].TextParts[0].Text = "New Content";
        }
    }
}
//Save the PowerPoint Presentation
pptxDoc.Save(Path.GetFullPath(@"Output/Result.pptx"));