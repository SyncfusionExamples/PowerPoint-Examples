using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a SmartArt to the slide at the specified size and position.
ISmartArt smartArt = slide.Shapes.AddSmartArt(SmartArtType.OrganizationChart, 0, 0, 640, 426.96);
//Traverse through all nodes of the SmartArt.
foreach (ISmartArtNode node in smartArt.Nodes)
{
    foreach(ISmartArtNode childnode in node.ChildNodes)
    {
        //Check if the node is assistant or not.
        if (childnode.IsAssistant)
            //Set the assistant node to false.
            childnode.IsAssistant = false;

    }
    
}
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);