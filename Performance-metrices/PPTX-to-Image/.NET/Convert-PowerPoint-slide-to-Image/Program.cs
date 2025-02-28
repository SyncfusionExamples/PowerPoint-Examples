using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

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
            string slideOutputFolder = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileName));
            Directory.CreateDirectory(slideOutputFolder);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open presentation
                using (FileStream inputStream = new(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (IPresentation pptxDoc = Presentation.Open(inputStream))
                    {
                        pptxDoc.PresentationRenderer = new PresentationRenderer();

                        // Convert each slide to an image
                        for (int i = 0; i < pptxDoc.Slides.Count; i++)
                        {
                            using (Stream stream = pptxDoc.Slides[i].ConvertToImage(ExportImageFormat.Jpeg))
                            {
                                string imagePath = Path.Combine(slideOutputFolder, $"Slide_{i + 1}.jpg");
                                using (FileStream fileStreamOutput = new(imagePath, FileMode.Create, FileAccess.Write))
                                {
                                    stream.CopyTo(fileStreamOutput);
                                }
                            }
                        }
                    }
                }
                stopwatch.Stop();
                Console.WriteLine($"{fileName} processed in {stopwatch.Elapsed.TotalSeconds} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}

