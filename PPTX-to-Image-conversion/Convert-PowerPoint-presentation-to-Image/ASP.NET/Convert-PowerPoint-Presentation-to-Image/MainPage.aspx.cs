using System;
using System.Web;
using System.IO;
using System.Web.Configuration;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.Drawing;

namespace Convert_PowerPoint_Presentation_to_Image
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
                //Initialize the PresentationRenderer.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Converts the first slide into image.
                using (Stream stream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Jpeg))
                {
                    //Create the output image file stream.
                    using (FileStream fileStreamOutput = File.Create(Server.MapPath("~/PPTXtoImage.Jpeg")))
                    {
                        //Copy the converted image stream into created output image stream
                        stream.CopyTo(fileStreamOutput);
                    }
                }
            }           
        }
    }
}