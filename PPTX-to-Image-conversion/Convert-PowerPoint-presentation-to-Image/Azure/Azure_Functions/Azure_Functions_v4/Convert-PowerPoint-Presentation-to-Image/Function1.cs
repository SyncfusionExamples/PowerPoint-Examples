using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Microsoft.Azure.WebJobs.Host;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;

namespace Convert_PowerPoint_Presentation_to_Image
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input PowerPoint document as stream from request.
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing PowerPoint document.
            using (IPresentation pptxDoc = Presentation.Open(stream))
            {
                //Initialize the PresentationRenderer to perform image conversion.
                pptxDoc.PresentationRenderer = new PresentationRenderer();
                //Convert PowerPoint slide to image as stream.
                Stream imageStream = pptxDoc.Slides[0].ConvertToImage(ExportImageFormat.Png);
                //Reset the stream position.
                imageStream.Position = 0;
                // Create a new memory stream.
                MemoryStream memoryStream = new MemoryStream();
                // Copy the contents of the image stream to the memory stream.
                imageStream.CopyTo(memoryStream);
                //Reset the stream position.
                memoryStream.Position = 0;
                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the image file saved stream as content of response.
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "PPTXtoImage.Jpeg"
                };
                //Set the content type as image mime type.
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/jpeg");
                //Return the response with output image stream.
                return response;                          
            }
        }
    }
}
