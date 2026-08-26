#  Split Presentation using C#

The [.NET PowerPoint library](https://www.syncfusion.com/document-sdk/net-powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **Split an PowerPoint Presentation** using C#.

## Steps to Split an PowerPoint Presentation programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.Compression.Zip;
```

Step 4: Add the following code snippet in Program.cs file to Split in the PowerPoint Presentation.

```csharp
ZipArchive zipArchive = new ZipArchive();
zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

// Open the source PowerPoint presentation.
IPresentation sourcePptx = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));

// Iterate through each section in the presentation.
foreach (ISection section in sourcePptx.Sections)
{
    // Create a new destination presentation for the current section.
    IPresentation destinationPptx = Presentation.Create();

    // Clone all slides from the current section and add them to the new presentation.
    foreach (ISlide slide in section.Slides)
    {
        destinationPptx.Slides.Add(slide.Clone(), PasteOptions.SourceFormatting, sourcePptx);
    }
    // Save the section presentation to a memory stream.
    MemoryStream memoryStream = new MemoryStream();
    destinationPptx.Save(memoryStream);

    // Add the generated presentation to the ZIP archive with the section name as the file name.
	string outputPath = Path.Combine(section.Name + "_Slides.pptx");
    zipArchive.AddItem(outputPath, memoryStream, true, Syncfusion.Compression.FileAttributes.Normal);

    // Close the destination presentation.
    destinationPptx.Close();
}

// Save the ZIP archive containing all section presentations.
zipArchive.Save(Path.GetFullPath(@"Output/Split-PowerPoint-by-sections.zip"));
// Close the ZIP archive.
zipArchive.Close();
// Close the source presentation.
sourcePptx.Close();
```

More information about create an PowerPoint Presentation, you can be refer in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-powerpoint-presentation) section.