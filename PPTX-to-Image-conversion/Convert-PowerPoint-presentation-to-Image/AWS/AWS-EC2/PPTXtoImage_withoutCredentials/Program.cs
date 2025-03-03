using Amazon.S3;
using Amazon;
using Amazon.S3.Model;
using System;
using System.IO;
using System.Threading.Tasks;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using SkiaSharp;

namespace PPTXtoImage
{
    class Program
    {
        private static readonly RegionEndpoint bucketRegion = RegionEndpoint.USEast1;
        private static IAmazonS3 s3Client;

        static async Task Main()
        {
            var config = new AmazonS3Config
            {
                RegionEndpoint = bucketRegion
            };
            s3Client = new AmazonS3Client(config);

            Console.WriteLine("Kindly enter the S3 bucket name: ");
            string bucketName = Console.ReadLine();
            Console.WriteLine("Kindly enter the input folder name that has the input PowerPoints: ");
            string inputFolderName = Console.ReadLine();
            Console.WriteLine("Kindly enter the output folder name in which the output images should be stored: ");
            string outputFolderName = Console.ReadLine();
            //Gets the list of imput files from the input folder.
            List<string> inputFileNames = await ListFilesAsync(inputFolderName, bucketName);

            for (int i = 0; i < inputFileNames.Count; i++)
            {
                //Converts PPTX to Image.
                await ConvertPptxToImage(inputFileNames[i], inputFolderName, bucketName, outputFolderName);
            }
        }
        /// <summary>
        /// Gets the list of input files.
        /// </summary>
        /// <returns></returns>
        private static async Task<List<string>> ListFilesAsync(string inputFolderName, string bucketName)
        {
            List<string> files = new List<string>();
            try
            {
                var request = new ListObjectsV2Request
                {
                    BucketName = bucketName,
                    Prefix = $"{inputFolderName}/",
                    Delimiter = "/"
                };

                ListObjectsV2Response response;
                do
                {
                    response = await s3Client.ListObjectsV2Async(request);
                    foreach (S3Object entry in response.S3Objects)
                    {
                        // Skip the "folder" itself
                        if (entry.Key.EndsWith("/"))
                            continue;

                        // Extract the filename by removing the folder path prefix
                        string fileName = entry.Key.Substring(inputFolderName.Length);
                        files.Add(fileName);
                    }
                    request.ContinuationToken = response.NextContinuationToken;
                } while (response.IsTruncated);
                return files;
            }
            catch (AmazonS3Exception e)
            {
                Console.WriteLine("Error encountered on server. Message:'{0}' when listing objects", e.Message);
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown encountered on server. Message:'{0}' when listing objects", e.Message);
                return null;
            }
        }
        /// <summary>
        /// Converts the PPTX to the image.
        /// </summary>
        /// <param name="inputFileName">Input file name</param>
        /// <returns></returns>
        static async Task ConvertPptxToImage(string inputFileName, string inputFolderName, string bucketName, string outputFolderName)
        {
            try
            {
                //Download the file from S3 into the MemoryStream
                var response = await s3Client.GetObjectAsync(new Amazon.S3.Model.GetObjectRequest
                {
                    BucketName = bucketName,
                    Key = inputFolderName+inputFileName
                });
                using (Stream responseStream = response.ResponseStream)
                {
                    MemoryStream fileStream = new MemoryStream();
                    await responseStream.CopyToAsync(fileStream);
                    //Open the existing PowerPoint presentation.
                    using (IPresentation pptxDoc = Presentation.Open(fileStream))
                    {
                        //Initialize PresentationRenderer.
                        pptxDoc.PresentationRenderer = new PresentationRenderer();
                        //Convert the PowerPoint presentation as image streams.
                        Stream[] images = pptxDoc.RenderAsImages(ExportImageFormat.Png);
                        //Gets the file name without extension.
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputFileName);
                        //Save the image streams to file.
                        for (int i = 0; i < images.Length; i++)
                        {
                            using (Stream stream = images[i])
                            {
                                //Uploads the image to the S3 bucket.
                                await UploadImageAsync(stream, $"{fileNameWithoutExt}" + i + ".png", bucketName, outputFolderName);
                            }
                        }
                    }
                }
            }
            catch (AmazonS3Exception e)
            {
                Console.WriteLine($"Error encountered on server. Message:'{e.Message}'");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Unknown error encountered. Message:'{e.Message}'");
            }
        }
        /// <summary>
        /// Uploads the image to the S3 bucket.
        /// </summary>
        /// <param name="imageStream">Converted image stream</param>
        /// <param name="outputFileName">Name that the image should be saved</param>
        /// <returns></returns>
        public static async Task UploadImageAsync(Stream imageStream, string outputFileName, string bucketName, string outputFolderName)
        {
            try
            {
                var key = $"{outputFolderName}/{outputFileName}"; // e.g., "images/your-image.png"

                var request = new PutObjectRequest
                {
                    BucketName = bucketName,
                    Key = key,
                    InputStream = imageStream,
                    ContentType = "image/png" // Adjust based on your image type
                };

                var response = await s3Client.PutObjectAsync(request);

                if (response.HttpStatusCode == System.Net.HttpStatusCode.OK)
                {
                    Console.WriteLine("Image uploaded successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to upload image. HTTP Status Code: {response.HttpStatusCode}");
                }
            }
            catch (AmazonS3Exception ex)
            {
                Console.WriteLine($"Error encountered on server. Message:'{ex.Message}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unknown error encountered. Message:'{ex.Message}'");
            }
        }
    }
}