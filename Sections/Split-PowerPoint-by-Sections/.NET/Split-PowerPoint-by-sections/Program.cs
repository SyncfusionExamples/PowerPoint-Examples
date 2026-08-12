// Create a ZIP archive and set the compression level to Best.
using Syncfusion.Presentation;
using Syncfusion.Compression.Zip;

ZipArchive zipArchive = new ZipArchive();
zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

// Open the source PowerPoint presentation.
IPresentation sourcePptx = Presentation.Open(Path.GetFullPath(@"Data/Template.pptx"));

// Iterate through each section in the presentation.
foreach (ISection section in sourcePptx.Sections)
{
    // Create a new destination presentation for the current section.
    IPresentation destinationPptx = Presentation.Create();

    // Clone all slides from the current section and add them to the new presentation.
    foreach (ISlide slide in section.Slides)
    {
        destinationPptx.Slides.Add(slide.Clone(), PasteOptions.SourceFormatting, sourcePptx);
    }
    // Save the section presentation to a memory stream.
    MemoryStream memoryStream = new MemoryStream();
    destinationPptx.Save(memoryStream);

    // Add the generated presentation to the ZIP archive with the section name as the file name.
	string outputPath = Path.Combine(section.Name + "_Slides.pptx");
    zipArchive.AddItem(outputPath, memoryStream, true, Syncfusion.Compression.FileAttributes.Normal);

    // Close the destination presentation.
    destinationPptx.Close();
}

// Save the ZIP archive containing all section presentations.
zipArchive.Save(Path.GetFullPath(@"Output/Split-PowerPoint-by-sections.zip"));
// Close the ZIP archive.
zipArchive.Close();
// Close the source presentation.
sourcePptx.Close();