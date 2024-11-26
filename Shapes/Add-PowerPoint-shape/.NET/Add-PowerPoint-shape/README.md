# Add shapes to a PowerPoint Presentation using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **add shapes to a PowerPoint Presentation** using C#.

## Steps to add shapes programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to add shapes to the PowerPoint Presentation.

```csharp
//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide.
slide.Shapes.AddShape(AutoShapeType.Cube, 50, 200, 300, 300);
//Get a picture as stream.
using FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.jpg"), FileMode.Open);
//Add picture to the shape collection.
IPicture picture = slide.Shapes.AddPicture(imageStream, 373, 83, 526, 382);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
```

More information about adding shapes can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-shapes) section.