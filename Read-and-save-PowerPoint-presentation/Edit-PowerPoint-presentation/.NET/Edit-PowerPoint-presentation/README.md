# Edit Presentation using C#

The [.NET PowerPoint library](https://www.syncfusion.com/document-sdk/net-powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can modify an existing PowerPoint presentation in C# — such as updating slide content, changing layouts, editing text, images, charts, and applying new styles.

## Steps to Edit an PowerPoint Presentation programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation; 
using System.IO; 
```

Step 4: Add the following code snippet in the Program.cs file to edit the PowerPoint presentation — update background, text, images, and table styles:

```csharp
//Opens the PPTX document. 
using IPresentation pptxDoc = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx")); 

//Get the new background picture as a stream. 
using FileStream bgPictureStream = new FileStream(Path.GetFullPath(@"Data/Background.png"), FileMode.Open); 
//Create an instance for memory stream 
using MemoryStream bgMemoryStream = new MemoryStream(); 
//Copy stream to memoryStream. 
bgPictureStream.CopyTo(bgMemoryStream); 

//Set the master slide background fill as Picture fill in PPTX document. 
pptxDoc.Masters[0].Background.Fill.FillType = FillType.Picture; 
pptxDoc.Masters[0].Background.Fill.PictureFill.ImageBytes = bgMemoryStream.ToArray(); 

//Get the first shape of the first slide from the PPTX document. 
IShape shape = pptxDoc.Slides[0].Shapes[0] as IShape; 
//Change the text of the shape. 
if (shape.TextBody.Text == "Company History") 
    shape.TextBody.Text = "Adventure Cycles History"; 

//Get the new picture as a stream. 
using FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/AdventureCycles.png"), FileMode.Open);
//Create an instance for memory stream 
using MemoryStream picMemoryStream = new MemoryStream(); 
//Copy stream to memoryStream. 
pictureStream.CopyTo(picMemoryStream); 

//Replace the existing image with the new image. 
pptxDoc.Slides[0].Pictures[0].ImageData = picMemoryStream.ToArray(); 

//Changes the built in style of the table. 
ITable table = pptxDoc.Slides[2].Tables[0]; 
table.BuiltInStyle = BuiltInTableStyle.ThemedStyle2Accent5; 

//Save the PPTX document 
pptxDoc.Save(Path.GetFullPath(@"Output/Output.pptx")); 
```

More information about create an PowerPoint Presentation, you can be refer in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-powerpoint-presentation) section.