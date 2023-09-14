using System;
using System.Web;

using System.IO;

using System.Web.Configuration;
using Syncfusion.Presentation;
using Syncfusion.Pdf;
using Syncfusion.PresentationRenderer;

namespace Convert_PowerPoint_Presentation_to_PDF
{
    public partial class WebForm1 : System.Web.UI.Page
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