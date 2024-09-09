using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Drawing;
using System.Drawing.Imaging;

namespace Convert_PowerPoint_Presentation_to_Image.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public void ConvertPPTXtoImage()
        {
            //Open the file as Stream.
            using (FileStream pathStream = new FileStream(Server.MapPath("~/App_Data/Input.pptx"), FileMode.Open, FileAccess.Read))
            {
                //Opens a PowerPoint Presentation.
                using (IPresentation pptxDoc = Presentation.Open(pathStream))
                {
                    //Convert the first slide into image.
                    Image image = pptxDoc.Slides[0].ConvertToImage(Syncfusion.Drawing.ImageType.Metafile);
                    //Saves the image file to MemoryStream.
                    MemoryStream stream = new MemoryStream();
                    //Download image file in the browser.
                    ExportAsImage(image, "PPTXToImage.Jpeg", ImageFormat.Jpeg, HttpContext.ApplicationInstance.Response);
                }
            }
        }
        //To download the image file.
        protected void ExportAsImage(Image image, string fileName, ImageFormat imageFormat, HttpResponse response)
        {
            if (ControllerContext == null)
                throw new ArgumentNullException("Context");
            string disposition = "content-disposition";
            response.AddHeader(disposition, "attachment; filename=" + fileName);
            if (imageFormat != ImageFormat.Emf)
                image.Save(response.OutputStream, imageFormat);
            Response.End();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}