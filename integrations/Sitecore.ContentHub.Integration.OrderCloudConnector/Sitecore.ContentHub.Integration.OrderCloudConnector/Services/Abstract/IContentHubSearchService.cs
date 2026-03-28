using Sitecore.ContentHub.Integration.OrderCloudConnector.Constants;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Framework.Essentials.LoadOptions;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    interface IContentHubSearchService
    {
        QueryNodeFilter GetRelationFilter(string relationName, RelationRole relationRole, IEnumerable<long> ids);

        Task<IEnumerable<IEntity>> SearchAfter(QueryNodeFilter filter, IEnumerable<Sorting>? sorting = null, IEntityLoadConfiguration? entityLoadConfiguration = null, int take = ApiConstants.ContentHubClient.SearchAfterDefaultTake);
    }
}
