# Find and replace text in a PowerPoint Presentation using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **find and replace text in a PowerPoint Presentation** using C#.

## Steps to find and replace text programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to find and replace text in the PowerPoint Presentation.

```csharp
//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Find all the occurrences of a particular text in the PowerPoint presentation.
ITextSelection[] textSelections = pptxDoc.FindAll("product", false, false);
foreach (ITextSelection textSelection in textSelections)
{
	//Get the found text as a single text part.
	ITextPart textPart = textSelection.GetAsOneTextPart();
	//Replace the text.
	textPart.Text = "Service";
}
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
```

More information about find and replace can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-find-and-replace) section.