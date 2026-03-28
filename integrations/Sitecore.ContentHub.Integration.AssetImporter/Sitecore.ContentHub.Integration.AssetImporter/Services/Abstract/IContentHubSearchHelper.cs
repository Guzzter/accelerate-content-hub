using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    internal interface IContentHubSearchHelper
    {
        Task<IEnumerable<IEntity>> SearchAfter(SearchAfterQuery query);
    }
}
