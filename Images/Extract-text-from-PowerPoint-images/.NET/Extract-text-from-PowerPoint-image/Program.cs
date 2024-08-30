using Syncfusion.OCRProcessor;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Presentation;
using System.Collections.Generic;
using System.IO;

namespace Extract_text_from_PowerPoint_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(@"../../../Template.pptx"))
            {
                List<MemoryStream> pictureStreamList = new List<MemoryStream>();
                //Retrieves the each slide from the Presentation.
                foreach (ISlide slide in pptxDoc.Slides) 
                {
                    //Retrieves all the picture from the slide.
                    IPictures pictures = slide.Pictures;
                    foreach (IPicture picture in pictures)
                    {
                        pictureStreamList.Add(new MemoryStream(picture.ImageData));
                    }
                }
                //Extract text from images using OCR processor.
                ExtractTextFromImages(pictureStreamList);
            }
        }

        /// <summary>
        /// Extracts text from images using OCR processor.
        /// </summary>
        /// <param name="pictureStreamList">List of picture stream.</param>
        private static void ExtractTextFromImages(List<MemoryStream> pictureStreamList)
        {
            //Inside bin folder, the tessdata folder contains the language data files.
            string tessdataPath = Path.GetFullPath(@"runtimes/tessdata");
            int i = 1;
            //Get each picture and extract its text.
            foreach (MemoryStream imgStream in pictureStreamList)
            {
                //Initialize the OCR processor by providing the path of the tesseract binaries.
                using (OCRProcessor processor = new OCRProcessor())
                {
                    //Set OCR language to process.
                    processor.Settings.Language = Languages.English;

                    //Sets Unicode font to preserve the Unicode characters in a PDF document.
                    FileStream fontStream = new FileStream(Path.GetFullPath("../../../ARIALUNI.ttf"), FileMode.Open);
                    processor.UnicodeFont = new PdfTrueTypeFont(fontStream, 8);
                    
                    //Perform the OCR process for an image stream.
                    string ocrText = processor.PerformOCR(imgStream, tessdataPath);

                    //Write the OCR'ed text in text file. 
                    using (StreamWriter writer = new StreamWriter(Path.GetFullPath(@"../../../OCRText_" + i + ".txt"), true))
                    {
                        writer.WriteLine(ocrText);
                    }
                }
                //Dispose the image streams.
                imgStream.Dispose();
                i++;
            }
        }
    }
}
