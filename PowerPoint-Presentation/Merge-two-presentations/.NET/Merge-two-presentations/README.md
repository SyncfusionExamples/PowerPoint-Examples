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
//Open an existing source presentation.
using FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourcePresentation.pptx"), FileMode.Open);
using IPresentation sourcePresentation = Presentation.Open(sourceStream);
//Open an existing destination presentation.
using FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationPresentation.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using IPresentation destinationPresentation = Presentation.Open(destinationStream);
//Iterate each slide in source and merge to presentation.
for (int i = 0; i < sourcePresentation.Slides.Count; i++)
{
    //Clone the slide of the source presentation.
    ISlide clonedSlide = sourcePresentation.Slides[i].Clone();
    //Merge the cloned slide into the destination presentation.
    destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme, sourcePresentation);
}
//Save the PowerPoint presentation.
using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create);
destinationPresentation.Save(outputStream);
```

More information about merging Presentations can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-slide#merging-slide) section.