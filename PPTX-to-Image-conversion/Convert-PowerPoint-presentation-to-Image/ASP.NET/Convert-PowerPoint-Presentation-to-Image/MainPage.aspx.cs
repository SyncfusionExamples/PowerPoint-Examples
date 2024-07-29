using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Drawing.Imaging;

namespace Convert_PowerPoint_Presentation_to_Image
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
                //Converts the first slide into image.
                System.Drawing.Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
                //Download image file in the browser.
                ExportAsImage(image, "PPTXToImage.Jpeg", ImageFormat.Jpeg, HttpContext.Current.Response);
            }
        }
        //To download the image file.
        protected void ExportAsImage(System.Drawing.Image image, string fileName, ImageFormat imageFormat, HttpResponse response)
        {
            string disposition = "content-disposition";
            response.AddHeader(disposition, "attachment; filename=" + fileName);
            if (imageFormat != ImageFormat.Emf)
                image.Save(response.OutputStream, imageFormat);
            Response.End();
        }
    }
}