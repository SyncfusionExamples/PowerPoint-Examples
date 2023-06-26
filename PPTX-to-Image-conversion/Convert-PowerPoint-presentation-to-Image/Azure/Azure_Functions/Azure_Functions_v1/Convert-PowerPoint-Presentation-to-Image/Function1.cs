using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Syncfusion.Presentation;
using Syncfusion.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Net.Http.Headers;

namespace Convert_PowerPoint_Presentation_to_Image
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input PowerPoint document as stream from request.
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing PowerPoint Presentation document.
            using (IPresentation pptxDoc = Presentation.Open(stream))
            {
                //Converts the first slide into image.
                System.Drawing.Image image = pptxDoc.Slides[0].ConvertToImage(ImageType.Metafile);
                //initializes a new instance of the MemoryStream.
                MemoryStream memoryStream = new MemoryStream();
                //Saves the image file.
                image.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Jpeg);
                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the image file saved stream as content of response.
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "PPTXtoImage.Jpeg"
                };
                //Set the content type as image file mime type.
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/jpeg");
                //Return the response with output image stream.
                return response;
            }
        }
    }
}
