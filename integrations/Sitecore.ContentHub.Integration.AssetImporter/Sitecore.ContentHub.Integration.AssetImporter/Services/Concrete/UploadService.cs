using Microsoft.Extensions.Logging;
using Microsoft.Net.Http.Headers;
using Sitecore.ContentHub.Integration.AssetImporter.Constants;
using Sitecore.ContentHub.Integration.AssetImporter.Models.ContentHub;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Sdk.WebClient.Http;
using System.Net.Mime;
using System.Text;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class UploadService(ILoggerFactory loggerFactory, IContentHubClientHelper contentHubClientHelper) : IUploadService
    {
        private readonly ILogger logger = loggerFactory.CreateLogger<UploadService>();
        private readonly IContentHubClientHelper contentHubClientHelper = contentHubClientHelper;

        public async Task<bool> Upload(string name, byte[] content)
        {
            logger.LogInformation($"Uploading {name}");

            var requestUploadResponse = await RequestUpload(name, content, ApiConstants.UploadConfiguration.AssetUpload, ApiConstants.UploadAction.NewAsset);

            await ExecuteUpload(requestUploadResponse.Url, name, content);

            var finaliseUploadResponse = await FinaliseUpload(requestUploadResponse.Content);

            if(finaliseUploadResponse.Success)
                logger.LogInformation($"Uploaded {name}. Id: {finaliseUploadResponse.AssetId}. Identifier: {finaliseUploadResponse.AssetIdentifier}");
            else
                logger.LogInformation($"Failed to upload {name}. Error: {finaliseUploadResponse.Message}");
            return finaliseUploadResponse.Success;
        }

        async Task<UploadRequestResponse> RequestUpload(string name, byte[] content, string uploadConfigurationName, string uploadActionName)
        {
            var uploadContent = $"{{\"file_name\": \"{name}\", \"file_size\": {content.Length}, \"upload_configuration\": {{ \"name\": \"{uploadConfigurationName}\" }}, \"action\": {{ \"name\": \"{uploadActionName}\" }} }}";
            var jsonContent = new StringContent(uploadContent, Encoding.UTF8, MediaTypeNames.Application.Json);
            var response = await contentHubClientHelper.Execute(client => client.Raw.PostAsync(ApiConstants.Endpoints.UploadRequest, jsonContent));


            response.EnsureSuccessStatusCode();
            var responseObject = await response.Content.ReadAsJsonAsync();

            var url = response.Headers.GetValues(HeaderNames.Location).First();

            return new UploadRequestResponse() { Url = url, Content = response.Content };
        }

        async Task ExecuteUpload(string url, string name, byte[] content)
        {
            if(content.Length > PlatformConstants.Upload.ChunkSize)
            {
                var chunks = SplitByteArray(content, PlatformConstants.Upload.ChunkSize);

                for (int i = 0; i < chunks.Length; i++)
                {
                    await ExecuteUploadChunk($"{url}&chunk={i}&chunks={chunks.Length}", name, chunks[i]);
                }
                await contentHubClientHelper.Execute(client => client.Raw.PostAsync($"{url}&chunks={chunks.Length}"));
            }
            else
                await ExecuteUploadChunk(url, name, content);
            
        }

        async Task ExecuteUploadChunk(string url, string name, byte[] content)
        {
            var formDataContent = new MultipartFormDataContent
            {
                { new ByteArrayContent(content), ApiConstants.FormDataKeys.UploadFile, name }
            };
            var response = await contentHubClientHelper.Execute(client => client.Raw.PostAsync(url, formDataContent));
            response.EnsureSuccessStatusCode();
        }

        async Task<FinaliseUploadResponse> FinaliseUpload(HttpContent content)
        {
            var response = await contentHubClientHelper.Execute(client => client.Raw.PostAsync(ApiConstants.Endpoints.UploadFinalise, content));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsJsonAsync<FinaliseUploadResponse>();
        }


        public static byte[][] SplitByteArray(byte[] byteArray, long chunkSize)
        {
            return [.. byteArray
                .Select((b, i) => new { b, i })
                .GroupBy(x => x.i / chunkSize)
                .Select(g => g.Select(x => x.b).ToArray())];
        }
    }
}
