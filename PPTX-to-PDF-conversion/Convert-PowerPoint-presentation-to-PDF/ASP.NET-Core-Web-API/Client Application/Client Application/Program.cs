class Program
{
    static async Task Main(string[] args)
    {
        // Create an HttpClient instance
        using (HttpClient client = new HttpClient())
        {
            try
            {
                // Send a GET request to a URL
                HttpResponseMessage response = await client.GetAsync("https://localhost:7241/api/Values/api/ConvertToPdf");

                // Check if the response is successful
                if (response.IsSuccessStatusCode)
                {
                    // Read the content as a string
                    Stream responseBody = await response.Content.ReadAsStreamAsync();
                    FileStream fileStream = File.Create("../../../Output/Output.pdf");
                    responseBody.CopyTo(fileStream);
                    fileStream.Close();
                }
                else
                {
                    Console.WriteLine("HTTP error status code: " + response.StatusCode);
                }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine("Request exception: " + e.Message);
            }
        }
    }
}
