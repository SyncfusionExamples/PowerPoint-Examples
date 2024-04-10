using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
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
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);