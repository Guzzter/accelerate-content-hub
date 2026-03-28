using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IContentHubQueryService
    {
        Task<long?> GetFirstIdAsync(Query query);

        Task<IEntity?> GetFirstAsync(Query query, IEntityLoadConfiguration? loadConfiguration = null);

        Task<IEnumerable<long>> QueryIdsAsync(Query query);

        Task<IEnumerable<IEntity>> QueryAsync(Query query, IEntityLoadConfiguration? loadConfiguration = null);
    }
}
