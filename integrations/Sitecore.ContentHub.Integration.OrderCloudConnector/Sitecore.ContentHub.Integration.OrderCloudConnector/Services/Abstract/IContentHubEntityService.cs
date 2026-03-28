using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    interface IContentHubEntityService
    {
        Task<IEntity> GetEntityById(long id, IEntityLoadConfiguration? loadConfiguration = null);

        Task<IList<IEntity>> GetEntityById(IEnumerable<long> ids, IEntityLoadConfiguration? loadConfiguration = null);

        Task<IEntity> GetEntityByIdentifier(string identifier, IEntityLoadConfiguration? loadConfiguration = null);

        Task<IList<IEntity>> GetEntityByIdentifier(IEnumerable<string> identifiers, IEntityLoadConfiguration? loadConfiguration = null);

        IEntityLoadConfiguration BuildLoadConfiguration(IEnumerable<string>? properties = null, IEnumerable<string>? relations = null);

        Task<T> GetPropertyValue<T>(IEntity entity, string propertyName);
    }
}
