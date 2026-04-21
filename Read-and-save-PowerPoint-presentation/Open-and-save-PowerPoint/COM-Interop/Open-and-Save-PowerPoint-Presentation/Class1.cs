using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Open_and_Save_PowerPoint_Presentation
{
    [ComVisible(true)]
    [Guid("6F6B4A3C-5C98-4D63-8C7A-5B2C0F1D9E11")] // Use a unique GUID
    [ProgId("Open_and_Save_PowerPoint_Presentation.Class1")] // Stable ProgID for VBScript
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Class1
    {
        public void OpenAndSavePowerPointPresentation(string inputFilePath, string outputFilePath)
        {
            // Open an existing PowerPoint presentation
            using (IPresentation pptxDoc = Presentation.Open(inputFilePath))
            {
                // Gets the first slide from the PowerPoint presentation
                ISlide slide = pptxDoc.Slides[0];
                // Gets the first shape of the slide
                IShape shape = slide.Shapes[0] as IShape;
                // Change the text of the shape
                if (shape.TextBody.Text == "Company History")
                    shape.TextBody.Text = "Company Profile";
                // Save the PowerPoint presentation
                pptxDoc.Save(outputFilePath);
            }
            
        }        
    }
}
