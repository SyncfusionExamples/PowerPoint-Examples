using Syncfusion.Presentation;
using System.ComponentModel;

//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"../../../Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Replace OLE with text.
ReplaceOLEWithText(pptxDoc, "Embedded file was here");
//Save output PowerPoint Presentation
using FileStream outputStream = new(Path.GetFullPath(@"../../../Data/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);


void ReplaceOLEWithText(IPresentation pptxDoc, string altText)
{
    double shapeLeft, shapeTop, shapeWidth, shapeHeignt = 0;
    //Iterate through the slides collection.
    for (int i = 0; i < pptxDoc.Slides.Count; i++)
    {
        for (int j = 0; j < pptxDoc.Slides[i].Shapes.Count; j++)
        {

            if (pptxDoc.Slides[i].Shapes[j] is IOleObject)
            {
                IOleObject ole = pptxDoc.Slides[i].Shapes[j] as IOleObject;
                shapeLeft = ole.Left;
                shapeTop = ole.Top;
                shapeWidth = ole.Width;
                shapeHeignt = ole.Height;
                pptxDoc.Slides[i].Shapes.Remove(ole);
                //Insert a textbox with the alternate text in the bounds of OLE.
                pptxDoc.Slides[i].AddTextBox(shapeLeft, shapeTop, shapeWidth, shapeHeignt).TextBody.AddParagraph(altText);
            }
        }
    }
}