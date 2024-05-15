using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using SharePointContentAccess.Models;
using Microsoft.Extensions.Options; // Ensure this using directive is present

using System.Diagnostics;
using Microsoft.Graph.Models;
using System.Text;


namespace SharePointContentAccess.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly AzureAdConfig _azureAdConfig;
        private static readonly List<string> sharePointSites = new List<string>
        {
            "/sites/LACServiceDeliveryWiki",
            "/sites/APACServiceDeliveryWiki",
            "/sites/EMEAServiceDeliveryWiki",
            "/sites/NAMServiceDeliveryWiki",
            "/sites/JAPANServiceDeliveryWiki"
        };

        public HomeController(ILogger<HomeController> logger, IOptions<AzureAdConfig> azureAdOptions)
        {
            _logger = logger;
            _azureAdConfig = azureAdOptions.Value;
        }

        private static GraphServiceClient GetGraphServiceClient(AzureAdConfig azureAdConfig)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var options = new ClientSecretCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            var clientSecretCredential = new ClientSecretCredential(
                azureAdConfig.TenantId, azureAdConfig.ClientId, azureAdConfig.ClientSecret, options);
            return new GraphServiceClient(clientSecretCredential, scopes);
        }

        public async Task<IActionResult> Index()
        {
            var siteHost = "speetengroup.sharepoint.com";
            var filePath = Path.GetTempFileName();
            var stringBuilder = new StringBuilder();
            try
            {
                var graphClient = GetGraphServiceClient(_azureAdConfig);
                foreach (var sitePath in sharePointSites)
                {   
                    var siteUrl = $"{siteHost}:{sitePath}";
                    var site = await graphClient.Sites[siteUrl].GetAsync();
                    stringBuilder.AppendLine($"Site: {site.DisplayName}");
                    var pages = await graphClient.Sites[site.Id].Pages.GraphSitePage.GetAsync();

                    foreach (var page in pages.Value)
                    {
                        stringBuilder.AppendLine($"\tPage: {page.Title}");
                        var result = await graphClient.Sites[site.Id].Pages[page.Id].GraphSitePage.GetAsync((requestConfiguration) =>
                        {
                            requestConfiguration.QueryParameters.Expand = new string[] { "canvasLayout" };
                        });
                        var webparts = result?.CanvasLayout?.HorizontalSections?.SelectMany(section => section.Columns)
                            .SelectMany(column => column.Webparts);


                        if (webparts != null)
                        {
                            foreach (var webpart in webparts)
                            {
                                if (webpart is TextWebPart textPart)
                                {
                                    stringBuilder.AppendLine($"\t\tContent: {textPart.InnerHtml}");
                                }
                            }
                        }

                    }

                }

                await System.IO.File.WriteAllTextAsync(filePath, stringBuilder.ToString()); // Write accumulated text to file
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"An error occurred: {ex.Message}";
                return View();
            }

            
            return File(System.IO.File.ReadAllBytes(filePath), "text/plain", "SiteContentDetails.txt"); // Return the file for download
        }

        public async Task<IActionResult> DownloadAllFiles()
        {
            var siteHost = "speetengroup.sharepoint.com";
            var filePath = Path.GetTempFileName();
            var stringBuilder = new StringBuilder();
            var downloader = new FileDownloader();
            try
            {
                var graphClient = GetGraphServiceClient(_azureAdConfig);
                foreach (var sitePath in sharePointSites)
                {
                    var siteUrl = $"{siteHost}:{sitePath}";
                    var site = await graphClient.Sites[siteUrl].GetAsync();
                    
                    stringBuilder.AppendLine($"Site: {site.DisplayName}");
                    
                    var driveCollectionResponse = await graphClient.Sites[site.Id].Drives.GetAsync();
                    
                    foreach (var drive in driveCollectionResponse.Value)
                    {
                        var rootItem = await graphClient.Drives[drive.Id].Root.GetAsync();
                        await TraverseDriveItems(graphClient, drive.Id, rootItem.Id, stringBuilder, downloader);
                    }
                }
                await System.IO.File.WriteAllTextAsync(filePath, stringBuilder.ToString());
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"An error occurred: {ex.Message}";
                return View();
            }
            return File(System.IO.File.ReadAllBytes(filePath), "text/plain", "SiteAttachmentDetails.txt");
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private async Task TraverseDriveItems(GraphServiceClient graphClient, string driveId, string itemId, StringBuilder stringBuilder, FileDownloader downloader)
        {
            var children = await graphClient.Drives[driveId].Items[itemId].Children
                .GetAsync();

            foreach (var item in children.Value)
            {
                if (item.Folder != null)
                {
                    await TraverseDriveItems(graphClient, driveId, item.Id, stringBuilder, downloader);
                }
                else
                {
                    if (item.File != null && item.AdditionalData.ContainsKey("@microsoft.graph.downloadUrl") )
                    {
                        string downloadUrl = item.AdditionalData["@microsoft.graph.downloadUrl"].ToString();
                        string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
                        string localFilePath = Path.Combine(downloadsPath, item.Name);

                        stringBuilder.AppendLine($"File Name: {item.Name}");
                        stringBuilder.AppendLine($"Download URL: {downloadUrl}");

                        // Download the file
                        //try
                        //{
                        //    await downloader.DownloadFileAsync(downloadUrl, localFilePath);
                        //    stringBuilder.AppendLine($"Downloaded to: {localFilePath}");
                        //}
                        //catch (Exception ex)
                        //{
                        //    stringBuilder.AppendLine($"Error downloading {item.Name}: {ex.Message}");
                        //}
                    }
                }
            }
        }
    }
}


