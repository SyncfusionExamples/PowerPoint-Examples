# Add transition to a PowerPoint Presentation using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **add transition to a PowerPoint Presentation** using C#.

## Steps to add transition programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to add transitions to the PowerPoint Presentation.

```csharp
//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to the presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add a shape to the slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Set the transition effect type .
slide.SlideTransition.TransitionEffect = TransitionEffect.Checkerboard;
//Set the transition effect options.
slide.SlideTransition.TransitionEffectOption = TransitionEffectOption.Across;
//Save the PowerPoint Presentation as stream.
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
```

More information about adding transitions can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/create-edit-slide-transitions-in-powerpoint-presentation-slides-cs-vb-net) section.