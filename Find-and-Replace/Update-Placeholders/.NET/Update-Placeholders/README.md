# update placeholder text or images in a PowerPoint Presentation using C#

The Syncfusion&reg; [.NET PowerPoint Library](https://www.syncfusion.com/document-processing/powerpoint-framework/net/powerpoint-library) (Presentation) enables you to create, read, and edit PowerPoint files programmatically without Microsoft Office or interop dependencies. Using this library, you can **update placeholder text or images in a PowerPoint Presentation** using C#.

## Steps to update placeholder text or images programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.Presentation.Net.Core](https://www.nuget.org/packages/Syncfusion.Presentation.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using System.Text.Json;
using Syncfusion.Presentation;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to update placeholder text or images in the PowerPoint Presentation.

```csharp

// Read customer data from JSON file
string jsonContent = File.ReadAllText(Path.GetFullPath(@"Data/customers.json"));
JsonDocument jsonDoc = JsonDocument.Parse(jsonContent);
JsonElement customersArray = jsonDoc.RootElement.GetProperty("customers");

// Loop through each customer in the JSON array
for (int customerIndex = 0; customerIndex < customersArray.GetArrayLength(); customerIndex++)
{
    JsonElement customer = customersArray[customerIndex];
    string customerName = customer.GetProperty("Name").GetString();
    string price = customer.GetProperty("Price").GetString();
    string photo = customer.GetProperty("Photo").GetString();

    // Load a fresh copy of the template for each customer
    using FileStream inputStream = new(Path.GetFullPath(@"Data/UpdatePlaceholders.pptx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

    using (IPresentation pptxDoc = Presentation.Open(inputStream))
    {
                  // Maintain placeholder keys and their replacement values as arrays
            string[] placeholders    = { "{CustomerName}", "{Price}" };
            string[] replacementValues = { customerName,   price     };

            // Iterate through the placeholder array and perform replacement in a single loop
            for (int index = 0; index < placeholders.Length; index++)
            {
                //Finds all the occurrences of a particular text in the PowerPoint presentation
                ITextSelection[] textSelections = presentation.FindAll(placeholders[index], false, false);
                foreach (ITextSelection textSelection in textSelections)
                {
                    //Gets the found text as a single text part
                    ITextPart textPart = textSelection.GetAsOneTextPart();
                    //Replaces the text
                    textPart.Text = replacementValues[index];
                }
            }
        // Iterate through all slides in the presentation
        foreach (ISlide slide in pptxDoc.Slides)
        {
            // Iterate through each shape in the slide
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                IShape shape = slide.Shapes[i] as IShape;

                // Check for picture placeholder with specific alt text
                if (shape.SlideItemType is SlideItemType.Placeholder &&
                    shape.PlaceholderFormat.Type == PlaceholderType.Picture &&
                    shape.Description == "{ProfileImage}")
                {
                    // Build the image path using the photo filename from JSON
                    string imagePath = Path.GetFullPath($@"Data/{photo}");

                    // Replace the image
                    FileStream pictureStream = new FileStream(imagePath, FileMode.Open);
                    MemoryStream memoryStream = new MemoryStream();
                    pictureStream.CopyTo(memoryStream);
                    slide.Shapes.AddPicture(memoryStream, shape.Left, shape.Top, shape.Width, shape.Height);
                    slide.Shapes.Remove(shape);

                    // Close the picture streams
                    pictureStream.Close();
                    memoryStream.Close();
                    break;
                }
            }
        }

        // Save each customer's output as a separate file
        string sanitizedName = customerName.Replace(" ", "_");
        using FileStream outputStream = new(Path.GetFullPath($@"Output/{sanitizedName}_Profile.pptx"), FileMode.Create, FileAccess.ReadWrite);

        // Save the modified presentation
        pptxDoc.Save(outputStream);
    }
}
```

More information about find and replace can be found in this [documentation](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/working-with-find-and-replace) section.