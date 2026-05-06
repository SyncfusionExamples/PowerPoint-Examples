using System.Text.Json;
using Syncfusion.Presentation;

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
        // Replace {CustomerName} placeholder
        ITextSelection[] textSelections1 = pptxDoc.FindAll("{CustomerName}", false, false);
        foreach (ITextSelection textSelection in textSelections1)
        {
            // Gets the found text as a single text part
            ITextPart textPart = textSelection.GetAsOneTextPart();
            // Replaces the text
            textPart.Text = customerName;
        }

        // Replace {Price} placeholder
        ITextSelection[] textSelections2 = pptxDoc.FindAll("{Price}", false, false);
        foreach (ITextSelection textSelection in textSelections2)
        {
            // Gets the found text as a single text part
            ITextPart textPart = textSelection.GetAsOneTextPart();
            // Replaces the text
            textPart.Text = price;
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