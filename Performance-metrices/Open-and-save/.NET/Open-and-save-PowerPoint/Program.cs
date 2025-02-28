using System.Diagnostics;
using Syncfusion.Presentation;

class Program
{
    static void Main()
    {
        string inputFolder = Path.GetFullPath("../../../Data/");
        string outputFolder = Path.GetFullPath("../../../Output/");

        Directory.CreateDirectory(outputFolder);

        // Get all .pptx files in the Data folder
        string[] files = Directory.GetFiles(inputFolder, "*.pptx");

        foreach (string inputPath in files)
        {
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputFolder, fileName);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Load or open PowerPoint Presentation
                using FileStream inputStream = new(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using IPresentation pptxDoc = Presentation.Open(inputStream);

                // Save the presentation to Output folder
                using FileStream outputStream = new(outputPath, FileMode.Create, FileAccess.ReadWrite);
                pptxDoc.Save(outputStream);

                stopwatch.Stop();
                Console.WriteLine($"{fileName} open and saved in {stopwatch.Elapsed.TotalSeconds} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}
