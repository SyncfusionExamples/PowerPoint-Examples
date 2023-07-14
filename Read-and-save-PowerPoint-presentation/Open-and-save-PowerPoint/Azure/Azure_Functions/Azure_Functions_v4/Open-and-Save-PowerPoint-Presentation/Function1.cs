using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Azure.WebJobs.Host;
using System.Net.Http;
using Syncfusion.Presentation;
using System.Net;
using System.Net.Http.Headers;

namespace Open_and_Save_PowerPoint_Presentation
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input PowerPoint document as stream from request
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Open an existing PowerPoint presentation
            using (IPresentation pptxDoc = Presentation.Open(stream))
            {
                //Gets the first slide from the PowerPoint presentation
                ISlide slide = pptxDoc.Slides[0];
                //Gets the first shape of the slide
                IShape shape = slide.Shapes[0] as IShape;
                //Change the text of the shape
                if (shape.TextBody.Text == "Company History")
                    shape.TextBody.Text = "Company Profile";
                MemoryStream memoryStream = new MemoryStream();
                pptxDoc.Save(memoryStream);
                //Reset the memory stream position.
                memoryStream.Position = 0;
                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                //Set the PowerPoint document saved stream as content of response.
                response.Content = new ByteArrayContent(memoryStream.ToArray());
                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Sample.pptx"
                };
                //Set the content type as PowerPoint document mime type.
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pptx");
                //Return the response with output PowerPoint document stream.
                return response;
            }
        }
    }
}
