//Opens the PPTX document.
using Syncfusion.Presentation;
using System.IO;

IPresentation pptxDoc = Presentation.Open(@"Data/Template.pptx");

//Get the new background picture as a stream.
FileStream bgPictureStream = new FileStream(@"Data/Background.png", FileMode.Open);
//Create an instance for memory stream
MemoryStream memoryStream = new MemoryStream();
//Copy stream to memoryStream.
bgPictureStream.CopyTo(memoryStream);
//Set the master slide background fill as Picture fill in PPTX document.
pptxDoc.Masters[0].Background.Fill.FillType = FillType.Picture;
pptxDoc.Masters[0].Background.Fill.PictureFill.ImageBytes = memoryStream.ToArray();

//Get the first shape of the first slide from the PPTX document.
IShape shape = pptxDoc.Slides[0].Shapes[0] as IShape;
//Change the text of the shape.
if (shape.TextBody.Text == "Company History")
    shape.TextBody.Text = "Adventure Cycles History";

//Get the new picture as a stream.
FileStream pictureStream = new FileStream(@"Data/AdventureCycles.png", FileMode.Open);
//Create an instance for memory stream
memoryStream = new MemoryStream();
//Copy stream to memoryStream.
pictureStream.CopyTo(memoryStream);
//Replace the existing image with the new image.
pptxDoc.Slides[0].Pictures[0].ImageData = memoryStream.ToArray();

//Changes the built in style of the table.
ITable table = pptxDoc.Slides[2].Tables[0];
table.BuiltInStyle = BuiltInTableStyle.ThemedStyle2Accent5;

//Save the PPTX document
pptxDoc.Save(@"Output/Output.pptx");
//Closes the PPTX document.
pptxDoc.Close();