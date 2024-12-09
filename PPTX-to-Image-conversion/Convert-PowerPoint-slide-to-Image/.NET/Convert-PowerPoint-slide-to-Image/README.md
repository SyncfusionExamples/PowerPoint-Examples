# Convert PowerPoint Presentation to Image using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, edit, and convert PowerPoint files programmatically without Microsoft office or interop dependencies. Using this library, you can **convert a PowerPoint Presentation to Image** using C#.

## Steps to convert PPTX to Image programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.PresentationRenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.PresentationRenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a PowerPoint Presentation to image.

```csharp
//Load or open an PowerPoint Presentation.
using (FileStream inputStream = new(Path.GetFullPath(@"Data/Template.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing PowerPoint presentation.
    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
        //Initialize the PresentationRenderer to perform image conversion.
        pptxDoc.PresentationRenderer = new PresentationRenderer();
        //Convert the PowerPoint slide as an image stream .
        using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
        {
            //Create the output image file stream.
            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Image.jpg")))
            {
                //Copy the converted image stream into created output stream.
                stream.CopyTo(fileStreamOutput);
            }  
        }      
    }       
}
```

More information about PPTX to Image conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/conversions/powerpoint-to-image/net/presentation-to-image) section.