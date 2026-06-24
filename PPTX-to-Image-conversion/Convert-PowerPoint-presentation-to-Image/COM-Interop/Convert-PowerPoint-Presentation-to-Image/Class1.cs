using Syncfusion.Drawing;
using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Convert_PowerPoint_Presentation_to_Image
{
    [ComVisible(true)]
    [Guid("6F6B4A3C-5C98-4D63-8C7A-5B2C0F1D9E11")] // Use a unique GUID
    [ProgId("Convert_PowerPoint_Presentation_to_Image.Class1")] // Stable ProgID for VBScript
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Class1
    {
        public void ConvertPowerPointPresentationtoImage(string inputFilePath, string outputFilePath)
        {
            //Open a PowerPoint Presentation.
            using (IPresentation pptxDoc = Presentation.Open(inputFilePath))
            {
                //Converts the first slide into image.
                System.Drawing.Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
                //Save the image as jpeg.
                image.Save(outputFilePath);
            }
        }
    }
}
