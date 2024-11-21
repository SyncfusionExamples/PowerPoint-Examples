# Convert PowerPoint Presentation to PDF using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, edit, and convert PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **convert a PowerPoint Presentation to PDF** using C#.

## Steps to convert PPTX to PDF programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.PresentationRenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a PowerPoint Presentation to PDF.

```csharp
//Load or open an PowerPoint Presentation.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Create the MemoryStream to save the converted PDF.
using MemoryStream pdfStream = new();
//Convert the PowerPoint document to PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc);
//Save the converted PDF document to MemoryStream.
pdfDocument.Save(pdfStream);
pdfStream.Position = 0;
//Create the output PDF file stream.
using FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/PPTXToPDF.pdf"));
//Copy the converted PDF stream into created output PDF stream.
pdfStream.CopyTo(fileStreamOutput);
```

More information about PPTX to PDF conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/conversions/powerpoint-to-pdf/net/presentation-to-pdf) section.