using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    static async Task Main(string[] args)
    {
        List<string> tokens = new List<string> { "ghp_XXXXXXXXXXXXXXXXXXXXXX" }; // Add your GitHub token

        Console.WriteLine("Enter the full path of the Excel file containing the search IDs:");
        string excelFilePath = Console.ReadLine();

        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine("Error: Excel file not found.");
            return;
        }

        Console.WriteLine("Enter the directory where the output Excel file should be saved:");
        string outputDirectory = Console.ReadLine();

        if (!Directory.Exists(outputDirectory))
        {
            Console.WriteLine("Error: Output directory does not exist.");
            return;
        }

        Console.WriteLine("Reading IDs from Excel...");
        var ids = ReadIdsFromExcel(excelFilePath);

        Console.WriteLine("Searching GitHub for IDs...");
        var allResults = new List<SearchResultWithId>();

        foreach (var id in ids)
        {
            Console.WriteLine($"🔍 Searching for ID: {id}");

            // Step 1: Search GitHub for the ID in repository code
            var searchResults = await SearchGitHub(id, tokens);

            if (searchResults.Count == 0)
            {
                allResults.Add(new SearchResultWithId { Id = id, Repository = "NULL", Path = "NULL", MatchFound = "No" });
                continue;
            }

            foreach (var result in searchResults)
            {
                // Step 2: Fetch file content using raw.githubusercontent.com
                string fileContent = await FetchFileContent(result.Repository, "main", result.Path);

                // Step 3: Check if file content contains the ID
                bool matchFound = !string.IsNullOrEmpty(fileContent) &&
                                  (fileContent.Contains($"\"{id}\"") || fileContent.Contains($"'{id}'") || fileContent.Contains(id));

                allResults.Add(new SearchResultWithId
                {
                    Id = id,
                    Repository = result.Repository,
                    Path = result.Path,
                    MatchFound = matchFound ? "Yes" : "No"
                });
            }
        }

        // Step 4: Write results to Excel
        Console.WriteLine("Writing results to Excel...");
        string outputFilePath = Path.Combine(outputDirectory, "GitHubSearchResults.xlsx");
        WriteResultsToExcel(allResults, outputFilePath);

        Console.WriteLine($"✅ Results saved to {outputFilePath}");
    }

    // ✅ Fetch raw file content using raw.githubusercontent.com
    static async Task<string> FetchFileContent(string repo, string branch, string filePath)
    {
        using (HttpClient client = new HttpClient())
        {
            string rawUrl = $"https://raw.githubusercontent.com/charlesschwab/{repo}/{branch}/{filePath}";
            client.DefaultRequestHeaders.Add("User-Agent", "CSharp-Console-App");

            HttpResponseMessage response = await client.GetAsync(rawUrl);
            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadAsStringAsync();
            }
            return "";
        }
    }

    // ✅ Search GitHub for code matching the ID
    static async Task<List<SearchResultWithId>> SearchGitHub(string query, List<string> tokens)
    {
        var allResults = new List<SearchResultWithId>();
        int tokenIndex = 0;

        using (HttpClient client = new HttpClient())
        {
            string currentToken = tokens[tokenIndex];
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", currentToken);
            client.DefaultRequestHeaders.Add("User-Agent", "CSharp-Console-App");

            string url = $"https://api.github.com/search/code?q={query}+user:charlesschwab&per_page=5&page=1";
            var response = await client.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                var searchResponse = JsonConvert.DeserializeObject<SearchResponse>(responseBody);

                if (searchResponse?.Items != null)
                {
                    foreach (var item in searchResponse.Items)
                    {
                        allResults.Add(new SearchResultWithId
                        {
                            Id = query,
                            Repository = item.Repository.FullName,
                            Path = item.Path
                        });
                    }
                }
            }
        }
        return allResults;
    }

    // ✅ Read search IDs from Excel
    static List<string> ReadIdsFromExcel(string filePath)
    {
        var ids = new List<string>();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                string id = worksheet.Cells[row, 1].Text.Trim();
                if (!string.IsNullOrWhiteSpace(id))
                {
                    ids.Add(id);
                }
            }
        }
        return ids;
    }

    // ✅ Write results to Excel
    static void WriteResultsToExcel(List<SearchResultWithId> results, string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Results");
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Repository";
            worksheet.Cells[1, 3].Value = "Path";
            worksheet.Cells[1, 4].Value = "Match Found";

            for (int i = 0; i < results.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = results[i].Id;
                worksheet.Cells[i + 2, 2].Value = results[i].Repository;
                worksheet.Cells[i + 2, 3].Value = results[i].Path;
                worksheet.Cells[i + 2, 4].Value = results[i].MatchFound;
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}

// ✅ Supporting Classes
class SearchResponse { [JsonProperty("items")] public List<SearchResult> Items { get; set; } }
class SearchResult { public Repository Repository { get; set; } public string Path { get; set; } }
class Repository { public string FullName { get; set; } }
class SearchResultWithId
{
    public string Id { get; set; }
    public string Repository { get; set; }
    public string Path { get; set; }
    public string MatchFound { get; set; }
}
