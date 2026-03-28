using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.Search;
using System.Globalization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IContentHubEntityService
    {
        Task<IEntity> GetEntityById(long id, IEntityLoadConfiguration? loadConfiguration = null);

        Task<IEntity> GetEntityByIdentifier(string identifier, IEntityLoadConfiguration? loadConfiguration = null);

        Task<T?> GetRawEntityById<T>(long id);

        Task<IEnumerable<IEntity>> GetRelatedEntities(IEnumerable<IEntity> entities, Relation relation, IEntityLoadConfiguration entityLoadConfiguration);

        Task<IEnumerable<IEntity>> GetRelatedEntities(IEntity entity, IEnumerable<Relation> relations, IEntityLoadConfiguration? entityLoadConfiguration = null);

        Task<IEnumerable<IEntity>> GetRelatedEntities(IEnumerable<IEntity> entities, IEnumerable<Relation> relations, IEntityLoadConfiguration? entityLoadConfiguration = null);

        IDictionary<string, object?> GetPropertyValueForCultures(IEntity entity, string propertyName, CultureInfo? fallbackCulture);

        IDictionary<string, object?> GetOptionListValueForCultures(IEntity entity, string propertyName, IDataSource optionList, CultureInfo? fallbackCulture);
    }
}
