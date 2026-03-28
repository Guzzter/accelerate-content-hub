using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class ContentHubAssetService(ILogger<ContentHubAssetService> logger, IContentHubClientHelper clientHelper, IContentHubEntityService entityService, IContentHubQueryBuilderService queryBuilderService, IContentHubQueryService queryService) : IContentHubAssetService
    {
        public async Task<string?> GetExtractedContent(long assetId)
        {
            var assetRawEntity = await entityService.GetRawEntityById<RawEntityResponse>(assetId);

            if (assetRawEntity == null || !assetRawEntity.Renditions.TryGetValue(DocumentConstants.ExtractedContentRenditionName, out var rendition))
                return null;

            var renditionResponse = await clientHelper.Execute(client => client.Raw.GetAsync(rendition.First().Uri));
            renditionResponse.EnsureSuccessStatusCode();
            return await renditionResponse.Content.ReadAsStringAsync();
        }

        public Task<string?> GetExtractedContent(IEntity assetEntity)
        {
            if (!assetEntity.Id.HasValue)
                throw new ArgumentException("Asset entity must have an Id.");
            return GetExtractedContent(assetEntity.Id.Value);
        }

        public async Task<string?> GetPublicLink(long assetId, string resourceName, bool createIfNotExists = false)
        {
            var query = queryBuilderService
                .AddDefinitionFilter(SchemaConstants.PublicLink.DefinitionName)
                .AddRelationFilter(SchemaConstants.PublicLink.Relations.AssetToPublicLink, assetId)
                .AddPropertyFilter(SchemaConstants.PublicLink.Properties.Resource, resourceName)
                .Build();

            var publicLinkEntityId = await queryService.GetFirstIdAsync(query);

            if (publicLinkEntityId == null && createIfNotExists)
            {
                logger.LogInformation($"Creating public link for asset {assetId} and resource {resourceName}");
                publicLinkEntityId = await CreatePublicLink(assetId, resourceName);
            }

            if(publicLinkEntityId == null || !publicLinkEntityId.HasValue)
            {
                logger.LogWarning($"Public link not found for asset {assetId} and resource {resourceName}");
                return null;
            }

            var publicLinkRawEntity = await entityService.GetRawEntityById<RawEntityResponse>(publicLinkEntityId.Value);
            return publicLinkRawEntity?.PublicLink;
        }

        public Task<string?> GetPublicLink(IEntity assetEntity, string resourceName, bool createIfNotExists = false)
        {
            if(!assetEntity.Id.HasValue)
                throw new ArgumentException("Asset entity must have an Id.");
            return GetPublicLink(assetEntity.Id.Value, resourceName, createIfNotExists);
        }

        public async Task<long> CreatePublicLink(long assetId, string rendition, DateTime? expirationDate = null, string? relativeUrl = null)
        {
            return await clientHelper.Execute(client => client.Assets.CreatePublicLinkAsync(assetId, rendition, expirationDate, relativeUrl));
        }
    }
}
