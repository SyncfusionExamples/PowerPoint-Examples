using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

//Creating new zip archive
Syncfusion.Compression.Zip.ZipArchive zipArchive = new Syncfusion.Compression.Zip.ZipArchive();
//You can use CompressionLevel to reduce the size of the file.
zipArchive.DefaultCompressionLevel = Syncfusion.Compression.CompressionLevel.Best;

//Get the input files from the folder
string folderName = @"../../../InputDocuments";
string[] inputFiles = Directory.GetFiles(folderName);
DirectoryInfo directoryInfo = new DirectoryInfo(folderName);
List<string> files = new List<string>();
FileInfo[] fileInfo = directoryInfo.GetFiles();
foreach (FileInfo fi in fileInfo)
{
    string name = Path.GetFileNameWithoutExtension(fi.Name);
    files.Add(name);
}

//Converts each Word documents to PDF documents
for (int i = 0; i < inputFiles.Length; i++)
{
    //PDF file name
    string outputFileName = files[i] + ".pdf";
    //Loads the PowerPoint presentation into stream.
    using (FileStream fileStreamInput = new FileStream(inputFiles[i], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        MemoryStream outputStream = ConvertPresentationToPDF(fileStreamInput);
        //Add the converted PDF file to zip
        zipArchive.AddItem(outputFileName, outputStream, true, Syncfusion.Compression.FileAttributes.Normal);
    }
}

//Zip file name and location
FileStream zipStream = new FileStream(@"../../../OutputPDFs.zip", FileMode.OpenOrCreate, FileAccess.ReadWrite);
zipArchive.Save(zipStream, true);
zipArchive.Close();

MemoryStream ConvertPresentationToPDF(FileStream fileStreamInput)
{
    //Open the existing PowerPoint presentation with loaded stream.
    using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
    {
        //Create the MemoryStream to save the converted PDF.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            //Convert the PowerPoint document to PDF document.
            using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
            {
                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;
            }
            //Create the output PDF file stream
            MemoryStream fileStreamOutput = new MemoryStream();
            //Copy the converted PDF stream into created output PDF stream
            pdfStream.CopyTo(fileStreamOutput);
            return fileStreamOutput;
        }
    }
}