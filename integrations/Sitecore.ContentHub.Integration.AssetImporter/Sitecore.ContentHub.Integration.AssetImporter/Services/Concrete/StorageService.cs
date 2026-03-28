using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.AssetImporter.Models.Options;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class StorageService(ILoggerFactory loggerFactory, IOptions<AzureOptions> azureOptions, BlobServiceClient blobServiceClient) : IStorageService
    {
        private readonly ILogger logger = loggerFactory.CreateLogger<StorageService>();
        private readonly BlobServiceClient blobServiceClient = blobServiceClient;

        public async Task<IEnumerable<string>> GetFilesAsync()
        {
            logger.LogInformation("Listing files in container");
            var client = await GetContainerClient();
            return client.GetBlobs().AsEnumerable().Select(x => x.Name);
        }

        public async Task<byte[]> GetFileContentsAsync(string fileName)
        {
            logger.LogInformation($"Getting content for file {fileName}");
            var blobClient = await GetBlobClient(fileName);
            var response = await blobClient.DownloadAsync();
            using var memoryStream = new MemoryStream();
            response.Value.Content.CopyTo(memoryStream);
            return memoryStream.ToArray();
        }

        public async Task RemoveFileAsync(string fileName)
        {
            logger.LogInformation($"Removing file {fileName}");
            var blobClient = await GetBlobClient(fileName);
            await blobClient.DeleteIfExistsAsync();
        }

        async Task<BlobContainerClient> GetContainerClient()
        {
            var containerClient = blobServiceClient.GetBlobContainerClient(azureOptions.Value.StorageContainerName);
            await containerClient.CreateIfNotExistsAsync();
            return containerClient;
        }

        async Task<BlobClient> GetBlobClient(string blobName)
        {
            var containerClient = await GetContainerClient();
            return containerClient.GetBlobClient(blobName);
        }
    }
}
