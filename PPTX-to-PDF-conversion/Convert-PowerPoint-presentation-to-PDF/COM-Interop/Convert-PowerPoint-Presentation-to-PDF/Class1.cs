using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Convert_PowerPoint_Presentation_to_PDF
{
    [ComVisible(true)]
    [Guid("6F6B4A3C-5C98-4D63-8C7A-5B2C0F1D9E11")] // Use a unique GUID
    [ProgId("Convert_PowerPoint_Presentation_to_PDF.Class1")] // Stable ProgID for VBScript
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Class1
    {
        public void ConvertPowerPointPresentationtoPDF(string inputFilePath, string outputFilePath)
        {
            using(IPresentation pptxDoc = Presentation.Open(inputFilePath))
            {
                //Converts the PowerPoint Presentation into PDF document
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    // Save PDF document
                    pdfDocument.Save(outputFilePath);
                }
            }
        }
    }
}
