# Update placeholder text or images in a PowerPoint Presentation using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **Update placeholder text or images in a PowerPoint Presentation** using C#.

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
using FileStream inputStream = new(Path.GetFullPath(@"Data/UpdatePlaceholders.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

using (IPresentation pptxDoc = Presentation.Open(inputStream))
{
    //Finds all the occurrences of a particular text in the PowerPoint presentation
    ITextSelection[] textSelections1 = pptxDoc.FindAll("{CustomerName}", false, false);
    foreach (ITextSelection textSelection in textSelections1)
    {
        //Gets the found text as a single text part
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text
        textPart.Text = "Nancy Dovolio";
    }
    ITextSelection[] textSelections2 = pptxDoc.FindAll("{Price}", false, false);
    foreach (ITextSelection textSelection in textSelections2)
    {
        //Gets the found text as a single text part
        ITextPart textPart = textSelection.GetAsOneTextPart();
        //Replaces the text
        textPart.Text = "$50";
    }
    // Iterate through all slides in the presentation
    foreach (ISlide slide in pptxDoc.Slides)
    {
        // Iterate through each shape in the slide
        for (int i = 0; i < slide.Shapes.Count; i++)
        {
            IShape shape = slide.Shapes[i] as IShape;

            // Check for picture with specific alt text
            if (shape.SlideItemType is SlideItemType.Placeholder && shape.PlaceholderFormat.Type == PlaceholderType.Picture)
            {
                // Check if the picture has the specific alt text
                if (shape.Description == "{ProfileImage}")
                {
                    // Replace the image
                    FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Customer_profile.png"), FileMode.Open);
                    MemoryStream memoryStream = new MemoryStream();
                    pictureStream.CopyTo(memoryStream);
                    slide.Shapes.AddPicture(memoryStream, shape.Left, shape.Top, shape.Width, shape.Height);
                    slide.Shapes.Remove(shape);
                    // Close the picture stream
                    pictureStream.Close();
                    memoryStream.Close();
                }
            }
        }
    }
    using FileStream outputStream = new(Path.GetFullPath(@"Output/Result.pptx"), FileMode.Create, FileAccess.ReadWrite);
    // Save the modified presentation
    pptxDoc.Save(outputStream);
```

More information about find and replace can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-find-and-replace) section.