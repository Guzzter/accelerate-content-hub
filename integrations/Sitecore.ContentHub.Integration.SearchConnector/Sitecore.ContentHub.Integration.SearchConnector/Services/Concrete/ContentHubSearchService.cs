using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class ContentHubSearchService(IContentHubClientHelper clientHelper) : IContentHubSearchService
    {
        public QueryNodeFilter GetRelationFilter(string relationName, RelationRole relationRole, IEnumerable<long> ids)
        {
            if(ids == null || !ids.Any())
                throw new ArgumentException("Ids list cannot be empty");

            var relationFilter = new RelationQueryFilter(relationName, relationRole == RelationRole.Child ? Stylelabs.M.Framework.Essentials.LoadOptions.RelationRole.Child : Stylelabs.M.Framework.Essentials.LoadOptions.RelationRole.Parent);
            QueryNodeFilter? filter = null;
            foreach (var entityId in ids)
            {
                if (filter == null)
                    filter = relationFilter.Id == entityId;
                else
                    filter = new LogicalQueryFilter(filter, relationFilter.Id == entityId, LogicalOperator.Or);
            }
            return filter!;
        }

        public async Task<IEnumerable<IEntity>> SearchAfter(QueryNodeFilter filter, IEnumerable<Sorting>? sorting = null, IEntityLoadConfiguration? entityLoadConfiguration = null, int take = ApiConstants.ContentHubClient.SearchAfterDefaultTake)
        {
            sorting = sorting ?? [new() { Field = SchemaConstants.SystemProperties.Identifier, Order = QuerySortOrder.Asc }];
            var query = new SearchAfterQuery(filter, sorting) { Take = take, LoadConfiguration = entityLoadConfiguration ?? EntityLoadConfiguration.Default };

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
