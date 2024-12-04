# Convert PowerPoint Presentation to PDF using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, edit, and convert PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **convert a PowerPoint Presentation to PDF** using C#.

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
//Open the PowerPoint file stream.
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load an existing PowerPoint Presentation.
    using (IPresentation pptxDoc = Presentation.Open(fileStream))
    {
        //Convert PowerPoint into PDF document.
        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
        {
            //Save the PDF file to file system.
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/PPTXToPDF.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                pdfDocument.Save(outputStream);
            }
        }
    }
}
```

More information about PPTX to PDF conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/conversions/powerpoint-to-pdf/net/presentation-to-pdf) section.