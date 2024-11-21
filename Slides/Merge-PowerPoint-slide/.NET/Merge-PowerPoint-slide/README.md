# Merge PowerPoint Presentations using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **merge PowerPoint Presentations** using C#.

## Steps to merge PowerPoint Presentations programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to merge PowerPoint Presentations.

```csharp
//Load or open an PowerPoint Presentation.
using FileStream sourcePresentationStream = new(Path.GetFullPath(@"Data/SourcePresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using FileStream destinationPresentationStream = new(Path.GetFullPath(@"Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation sourcePresentation = Presentation.Open(sourcePresentationStream);
//Open the destination Presentation.
using IPresentation destinationPresentation = Presentation.Open(destinationPresentationStream);
//Clone the first slide of the source Presentation.
ISlide clonedSlide = sourcePresentation.Slides[0].Clone();
//Merge the cloned slide to the destination Presentation with paste option - Destination Theme.
destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme);
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
destinationPresentation.Save(outputStream);
```

More information about merging Presentations can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-slide#merging-slide) section.