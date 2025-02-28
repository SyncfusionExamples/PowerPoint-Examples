using System.Diagnostics;
using Syncfusion.Presentation;

class Program
{
    static void Main()
    {
        string sourceFolder = Path.GetFullPath("../../../Data/SourceDocument/");
        string destinationFolder = Path.GetFullPath("../../../Data/DestinationDocument/");
        string outputFolder = Path.GetFullPath("../../../Output/");

        Directory.CreateDirectory(outputFolder);

        // Get all source files
        string[] sourceFiles = Directory.GetFiles(sourceFolder, "*.pptx");

        foreach (string sourcePath in sourceFiles)
        {
            string fileName = Path.GetFileName(sourcePath);
            string destinationPath = Path.Combine(destinationFolder, fileName);

            if (!File.Exists(destinationPath))
            {
                Console.WriteLine($"Skipping {fileName} - No matching destination file found.");
                continue;
            }

            string outputPath = Path.Combine(outputFolder, fileName);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open source and destination presentations
                using FileStream sourceStream = new(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using IPresentation sourcePresentation = Presentation.Open(sourceStream);

                using FileStream destinationStream = new(destinationPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using IPresentation destinationPresentation = Presentation.Open(destinationStream);

                // Clone and merge all slides
                foreach (ISlide slide in sourcePresentation.Slides)
                {
                    ISlide clonedSlide = slide.Clone();
                    destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme);
                }

                // Save the merged presentation
                using FileStream outputStream = new(outputPath, FileMode.Create, FileAccess.ReadWrite);
                destinationPresentation.Save(outputStream);

                stopwatch.Stop();
                Console.WriteLine($"{fileName} is cloned and merged in {stopwatch.Elapsed.TotalSeconds} seconds.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}
