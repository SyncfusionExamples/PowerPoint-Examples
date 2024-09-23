using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using File = Google.Apis.Drive.v3.Data.File;
using Syncfusion.Presentation;

namespace Save_PowerPoint_document
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Create a new instance of PowerPoint Presentation file.
            IPresentation pptxDocument = Presentation.Create();

            //Add a new slide to file and apply background color.
            ISlide slide = pptxDocument.Slides.Add(SlideLayoutType.TitleOnly);
            //Specify the fill type and fill color for the slide background.
            slide.Background.Fill.FillType = FillType.Solid;
            slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(232, 241, 229);

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide.
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph("Company History").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Add description content to the slide by adding a new TextBox.
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = "IMN Solutions PVT LTD is the software company, established in 1987, by George Milton. The company has been listed as the trusted partner for many high-profile organizations since 1988 and got awards for quality products from reputed organizations.";

            //Add bullet points to the slide.
            IShape bulletPointsShape = slide.AddTextBox(53.22, 270, 437.90, 116.32);

            //Add a paragraph for a bullet point.
            IParagraph firstPara = bulletPointsShape.TextBody.AddParagraph("The company acquired the MCY corporation for 20 billion dollars and became the top revenue maker for the year 2015.");
            //Format how the bullets should be displayed.
            firstPara.ListFormat.Type = ListType.Bulleted;
            firstPara.LeftIndent = 35;
            firstPara.FirstLineIndent = -35;

            //Add another paragraph for the next bullet point.
            IParagraph secondPara = bulletPointsShape.TextBody.AddParagraph("The company is participating in top open source projects in automation industry.");
            //Format how the bullets should be displayed.
            secondPara.ListFormat.Type = ListType.Bulleted;
            secondPara.LeftIndent = 35;
            secondPara.FirstLineIndent = -35;

            //Gets a picture as stream.
            FileStream pictureStream = new FileStream(Path.GetFullPath("../../../Data/Image.jpg"), FileMode.Open);
            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Add an auto-shape to the slide.
            IShape stampShape = slide.Shapes.AddShape(AutoShapeType.Explosion1, 48.93, 430.71, 104.13, 80.54);
            //Format the auto-shape color by setting the fill type and text.
            stampShape.Fill.FillType = FillType.None;
            stampShape.TextBody.AddParagraph("IMN").HorizontalAlignment = HorizontalAlignmentType.Center;

            //Saves the PowerPoint to MemoryStream.
            MemoryStream stream = new MemoryStream();
            pptxDocument.Save(stream);

            // Load Google Drive API credentials from a file.
            UserCredential credential;
            string[] Scopes = { DriveService.Scope.Drive };
            string ApplicationName = "YourAppName";

            using (var stream1 = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))//Replace with your actual credentials.json
            {
                string credPath = "token.json";
                // Authorize the Google Drive API access.
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream1).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }
            // Create a new instance of Google Drive service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Create metadata for the file to be uploaded.
            var fileMetadata = new File()
            {
                Name = "Output.pptx", // Name of the file in Google Drive.
                MimeType = "application/powerpoint",
            };
            FilesResource.CreateMediaUpload request;
            // Create a memory stream from the PowerPoint presentation.
            using (var fs = new MemoryStream(stream.ToArray()))
            {
                // Create an upload request for Google Drive.
                request = service.Files.Create(fileMetadata, fs, "application/powerpoint");
                // Upload the file.
                request.Upload();
            }
            //Dispose the memory stream.
            stream.Dispose();
            pptxDocument.Close();
        }
    }
}
