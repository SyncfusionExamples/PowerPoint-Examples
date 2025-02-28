using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;

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
            string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileName) + ".pdf");

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open and convert presentation to PDF
                using (FileStream inputStream = new(inputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (IPresentation pptxDoc = Presentation.Open(inputStream))
                    {
                        pptxDoc.PresentationRenderer = new PresentationRenderer();

                        // Convert PowerPoint to PDF
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                        {
                            using (FileStream outputStream = new(outputPath, FileMode.Create, FileAccess.Write))
                            {
                                pdfDocument.Save(outputStream);
                            }
                        }
                    }
                }
                stopwatch.Stop();
                Console.WriteLine($"{fileName} taken time to convert as PDF: {stopwatch.Elapsed.TotalSeconds} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}

