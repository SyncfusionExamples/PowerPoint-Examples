# Format a table in a PowerPoint Presentation using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **format a table in a PowerPoint Presentation** using C#.

## Steps to format a table programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.Presentation;
using Syncfusion.Drawing;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to format a table in the PowerPoint Presentation.

```csharp
//Open an existing PowerPoint presentation.
using IPresentation pptxDoc = Presentation.Create();
//Add slide to the Presentation
ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
//Add table to the slide.
ITable table = slide.Shapes.AddTable(2, 2, 100, 120, 300, 200);
//Retrieve each cell and fills text content to the cell.

ICell cell = table[0, 0];
//Set the column width for a cell; this sets the width for entire column.
cell.ColumnWidth = 400;
//Set the margin for the cell.
cell.TextBody.MarginBottom = 0;
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.Orange;
cell.TextBody.AddParagraph("First Row and First Column");

cell = table[0, 1];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.BlueViolet;
cell.TextBody.AddParagraph("First Row and Second Column");

cell = table[1, 0];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.SandyBrown;
cell.TextBody.AddParagraph("Second Row and First Column");

cell = table[1, 1];
//Set the margin for the cell.
cell.TextBody.MarginLeft = 58;
cell.TextBody.MarginRight = 29;
cell.TextBody.MarginTop = 65;
//Set the back color for the cell.
cell.Fill.SolidFill.Color = ColorObject.Silver;
cell.TextBody.AddParagraph("Second Row and Second Column");
using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
pptxDoc.Save(outputStream);
```

More information about table formatting can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-tables#applying-table-formatting) section.