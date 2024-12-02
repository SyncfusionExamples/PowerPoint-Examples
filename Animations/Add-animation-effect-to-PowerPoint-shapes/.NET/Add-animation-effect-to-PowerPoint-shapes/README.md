# Add animation to a PowerPoint Presentation using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **add animation to a PowerPoint Presentation** using C#.

## Steps to add animation programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to add animations to the PowerPoint Presentation.

```csharp
//Load or open an PowerPoint Presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add a blank slide to Presentation.
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add normal shape to slide.
IShape cubeShape = slide.Shapes.AddShape(AutoShapeType.Cube, 100, 100, 300, 300);
//Access the animation sequence to create effects.
ISequence sequence = slide.Timeline.MainSequence;
//Add bounce effect to the shape.
IEffect bounceEffect = sequence.AddEffect(cubeShape, EffectType.Bounce, EffectSubtype.None, EffectTriggerType.OnClick);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
```

More information about adding animations can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-animation) section.