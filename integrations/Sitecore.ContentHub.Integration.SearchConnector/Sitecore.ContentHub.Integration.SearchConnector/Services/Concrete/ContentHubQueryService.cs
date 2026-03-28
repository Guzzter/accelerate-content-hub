using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class ContentHubQueryService(IContentHubClientHelper clientHelper) : IContentHubQueryService
    {
        public async Task<long?> GetFirstIdAsync(Query query)
        {
            query.Take = 1;
            return (await QueryIdsAsync(query)).FirstOrDefault();
        }

        public async Task<IEntity?> GetFirstAsync(Query query, IEntityLoadConfiguration? loadConfiguration = null)
        {
            query.Take = 1;
            return (await QueryAsync(query, loadConfiguration)).FirstOrDefault();
        }

        public async Task<IEnumerable<long>> QueryIdsAsync(Query query)
        {
            var result = await clientHelper.Execute(client => client.Querying.QueryIdsAsync(query));
            return result.Items;
        }

        public async Task<IEnumerable<IEntity>> QueryAsync(Query query, IEntityLoadConfiguration? loadConfiguration = null)
        {
            var result = await clientHelper.Execute(client => client.Querying.QueryAsync(query, loadConfiguration));
            return result.Items;
        }
    }
}
