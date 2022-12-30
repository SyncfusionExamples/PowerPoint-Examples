using Syncfusion.Presentation;

//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Access the first master slide in PowerPoint file.
IMasterSlide masterSlide = pptxDoc.Masters[0];
//Retrieve the background instance.
IBackground background = masterSlide.Background;
//Set the fill type for background as Solid fill.
background.Fill.FillType = FillType.Solid;
//Get the instance for solid Fill.
ISolidFill solidFill = background.Fill.SolidFill;
//Set the color for solid fill object.
solidFill.Color = ColorObject.Green; 
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"../../../Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);