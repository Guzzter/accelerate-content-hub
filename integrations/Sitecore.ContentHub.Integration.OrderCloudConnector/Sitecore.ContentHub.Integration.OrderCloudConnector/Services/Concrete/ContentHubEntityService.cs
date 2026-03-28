using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Framework.Essentials.LoadOptions;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Concrete
{
    class ContentHubEntityService(IContentHubClientHelper clientHelper) : IContentHubEntityService
    {
        public Task<IEntity> GetEntityById(long id, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetAsync(id, loadConfiguration));
        }

        public Task<IList<IEntity>> GetEntityById(IEnumerable<long> ids, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetManyAsync(ids, loadConfiguration));
        }

        public Task<IEntity> GetEntityByIdentifier(string identifier, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetAsync(identifier, loadConfiguration));
        }

        public Task<IList<IEntity>> GetEntityByIdentifier(IEnumerable<string> identifiers, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetManyAsync(identifiers, loadConfiguration));
        }

        public IEntityLoadConfiguration BuildLoadConfiguration(IEnumerable<string>? properties = null, IEnumerable<string>? relations = null)
        {
            return new EntityLoadConfigurationBuilder()
                .WithProperties(properties ?? [])
                .WithRelations(relations ?? [])
                .InCultures(LoadOption.Default)
                .Build();
        }

        public async Task<T> GetPropertyValue<T>(IEntity entity, string propertyName)
        {
            var property = entity.GetProperty(propertyName) ?? throw new ArgumentException($"Property '{propertyName}' not found in entity '{entity.Identifier}'.");
            if ((property.IsMultiLanguage))
            {
                var culture = await clientHelper.Execute(c => c.Cultures.GetDefaultCultureCachedAsync());
                return entity.GetPropertyValue<T>(propertyName, culture);
            }
            return entity.GetPropertyValue<T>(propertyName);
        }
    }
}
