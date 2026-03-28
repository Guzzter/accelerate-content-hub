using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ContentHubSearchHelper(IContentHubClientHelper clientHelper) : IContentHubSearchHelper
    {

        public async Task<IEnumerable<IEntity>> SearchAfter(SearchAfterQuery query)
        {
            var entities = new List<IEntity>();
            var searchAfter = Enumerable.Empty<string>();
            while (true)
            {
                query.SearchAfter = searchAfter;
                var result = await clientHelper.Execute(c => c.Querying.SearchAfterAsync(query));
                if (result.ReturnedItems == 0)
                    break;
                entities.AddRange(result.Items);
                searchAfter = result.LastHitData;
            }
            return entities;
        }
    }
}
