using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IContentHubAssetService
    {
        Task<string?> GetExtractedContent(long assetId);

        Task<string?> GetExtractedContent(IEntity assetEntity);

        Task<string?> GetPublicLink(long assetId, string resourceName, bool createIfNotExists = false);

        Task<string?> GetPublicLink(IEntity assetEntity, string resourceName, bool createIfNotExists = false);
    }
}
