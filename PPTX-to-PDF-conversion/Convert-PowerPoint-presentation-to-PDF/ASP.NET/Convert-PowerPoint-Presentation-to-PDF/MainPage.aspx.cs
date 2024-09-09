using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Convert_PowerPoint_Presentation_to_PDF
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            string filePath = Server.MapPath("~/App_Data/Input.pptx");
            //Open the existing PowerPoint presentation.
            using (IPresentation pptxDoc = Presentation.Open(filePath))
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
                    using (FileStream fileStreamOutput = File.Create(Server.MapPath("~/Sample.pdf")))
                    {
                        //Copy the converted PDF stream into created output PDF stream
                        pdfStream.CopyTo(fileStreamOutput);
                    }
                }
            }
        }
    }
}