# Convert PowerPoint Presentation to PDF/A using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, edit, and convert PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **convert a PowerPoint Presentation to PDF/A** using C#.

## Steps to convert PPTX to PDF/A programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.PresentationRenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a PowerPoint Presentation to PDF/A.

```csharp
//Create a file stream to read the PowerPoint presentation file.
using FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Open(inputStream);
//Initialize the conversion settings.
PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
//Set the Pdf conformance level to A1B
pdfConverterSettings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
//Convert the PowerPoint presentation to PDF file.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings);
//Create new instance of file stream.
using FileStream pdfStream = new(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create);
//Save the generated PDF to file stream.
pdfDocument.Save(pdfStream);
```

More information about converting PPTX to PDF with conformance can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/conversions/powerpoint-to-pdf/net/presentation-to-pdf#pdf-conformance) section.