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
                HttpResponseMessage response = await client.GetAsync("https://localhost:7260/api/Values/api/PPTXToImage");

                // Check if the response is successful
                if (response.IsSuccessStatusCode)
                {
                    // Read the content as a string
                    Stream responseBody = await response.Content.ReadAsStreamAsync();
                    FileStream fileStream = File.Create("../../../Output/Output.jpeg");
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