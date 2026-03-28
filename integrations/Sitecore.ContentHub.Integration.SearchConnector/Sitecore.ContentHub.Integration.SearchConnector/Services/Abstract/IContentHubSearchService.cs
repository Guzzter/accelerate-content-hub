using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IContentHubSearchService
    {

        QueryNodeFilter GetRelationFilter(string relationName, RelationRole relationRole, IEnumerable<long> ids);

        Task<IEnumerable<IEntity>> SearchAfter(QueryNodeFilter filter, IEnumerable<Sorting>? sorting = null, IEntityLoadConfiguration? entityLoadConfiguration = null, int take = ApiConstants.ContentHubClient.SearchAfterDefaultTake);
    }
}
