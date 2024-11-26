# Convert PowerPoint Presentation to PDF/UA using C#

The Syncfusion [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, edit, and convert PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **convert a PowerPoint Presentation to PDF/UA** using C#.

## Steps to convert PPTX to PDF/UA programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.PresentationRenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a PowerPoint Presentation to PDF/UA.

```csharp
//Load the PowerPoint presentation into a stream.
using FileStream fileStreamInput = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
//Open the existing PowerPoint presentation with the loaded stream.
using IPresentation pptxDoc = Presentation.Open(fileStreamInput) ;
//Instantiation of the PresentationToPdfConverterSettings.
PresentationToPdfConverterSettings pdfConverterSettings = new PresentationToPdfConverterSettings();
//Enable a flag to preserve structured document tags in the converted PDF document.               
pdfConverterSettings.AutoTag = true;
//Convert the PowerPoint document to a PDF document.
using PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, pdfConverterSettings);
//Create new instance of file stream.
using FileStream fileStreamOutput = new(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create);
//Save the generated PDF to file stream.
pdfDocument.Save(fileStreamOutput);
```

More information about PPTX to PDF/UA conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/conversions/powerpoint-to-pdf/net/presentation-to-pdf#accessible-pdf-document) section.