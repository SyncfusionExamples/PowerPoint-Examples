using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.Net.Http;
using System.Net;
using Microsoft.Azure.WebJobs.Host;
using System.Net.Http.Headers;

namespace Convert_PowerPoint_Presentation_to_PDF
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            //Gets the input PowerPoint document as stream from request.
            Stream stream = req.Content.ReadAsStreamAsync().Result;
            //Loads an existing PowerPoint document
            using (IPresentation pptxDoc = Presentation.Open(stream))
            {
                //Creates an instance of the DocToPDFConverter.
                using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                {
                    MemoryStream memoryStream = new MemoryStream();
                    //Saves the PDF file .
                    pdfDocument.Save(memoryStream);
                    //Reset the memory stream position.
                    memoryStream.Position = 0;
                    //Create the response to return.
                    HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                    //Set the PDF document saved stream as content of response.
                    response.Content = new ByteArrayContent(memoryStream.ToArray());
                    //Set the contentDisposition as attachment.
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = "Sample.Pdf"
                    };
                    //Set the content type as PDF document mime type.
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");
                    //Return the response with output PDF document stream.
                    return response;
                                  
                }
            }
        }
    }
}
