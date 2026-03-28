using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using System.Globalization;
using System.Net.Http.Json;
using System.Text.Json;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class ContentHubEntityService(IContentHubClientHelper clientHelper, IContentHubSearchService searchService) : IContentHubEntityService
    {
        public Task<IEntity> GetEntityById(long id, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetAsync(id, loadConfiguration));
        }

        public Task<IEntity> GetEntityByIdentifier(string identifier, IEntityLoadConfiguration? loadConfiguration = null)
        {
            return clientHelper.Execute(client => client.Entities.GetAsync(identifier, loadConfiguration));
        }

        public IDictionary<string, object?> GetPropertyValueForCultures(IEntity entity, string propertyName, CultureInfo? fallbackCulture)
        {
            var property = entity.GetProperty(propertyName);
            return entity.Cultures.ToDictionary(c => c.Name, c => property.IsMultiLanguage ? GetPropertyValueForCulture(entity, propertyName, c, fallbackCulture) : entity.GetPropertyValue(propertyName));
        }

        public IDictionary<string, object?> GetOptionListValueForCultures(IEntity entity, string propertyName, IDataSource optionList, CultureInfo? fallbackCulture)
        {
            if (entity.GetProperty(propertyName).DataType.IsArray)
                return entity.Cultures.ToDictionary(c => c.Name, c => GetOptionListValueForCulture(entity.GetPropertyValue<string[]>(propertyName), optionList.GetDataSourceValues(), c, fallbackCulture));
            return entity.Cultures.ToDictionary(c => c.Name, c => GetOptionListValueForCulture(entity.GetPropertyValue<string>(propertyName), optionList.GetDataSourceValues(), c, fallbackCulture));
        }

        public async Task<T?> GetRawEntityById<T>(long id)
        {
            var response = await clientHelper.Execute(client => client.Raw.GetAsync(GetEntityUrl(id)));
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<T>(new JsonSerializerOptions() {  PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower });
        }

        public async Task<IEnumerable<IEntity>> GetRelatedEntities(IEnumerable<IEntity> entities, Relation relation, IEntityLoadConfiguration entityLoadConfiguration)
        {
            var filter = searchService.GetRelationFilter(relation.Name, relation.Role, entities.Select(x => x.Id!.Value));
            return await searchService.SearchAfter(filter, entityLoadConfiguration: entityLoadConfiguration);
        }

        public async Task<IEnumerable<IEntity>> GetRelatedEntities(IEntity entity, IEnumerable<Relation> relations, IEntityLoadConfiguration? entityLoadConfiguration = null)
        {
            if(!entity.Id.HasValue)
                return [];
            var entities = new List<IEntity> { entity }.AsEnumerable();
            return await GetRelatedEntities(entities, relations, entityLoadConfiguration);
        }

        public async Task<IEnumerable<IEntity>> GetRelatedEntities(IEnumerable<IEntity> entities, IEnumerable<Relation> relations, IEntityLoadConfiguration? entityLoadConfiguration = null)
        {
            foreach (var relation in relations)
                entities = await GetRelatedEntities(entities, relation, relations.Last() == relation ? entityLoadConfiguration ?? EntityLoadConfiguration.Default : EntityLoadConfiguration.Minimal);
            return entities;
        }

        private object? GetPropertyValueForCulture(IEntity entity, string propertyName, CultureInfo culture, CultureInfo? fallbackCulture)
        {
            return entity.GetPropertyValue(propertyName, culture) ?? (fallbackCulture != null ? entity.GetPropertyValue(propertyName, fallbackCulture) : null);
        }

        private object? GetOptionListValueForCulture(IEnumerable<string> identifier, IEnumerable<IDataSourceValue> dataSourceValues, CultureInfo culture, CultureInfo? fallbackCulture)
        {
            return identifier.Select(i => GetOptionListValueForCulture(i, dataSourceValues, culture, fallbackCulture));
        }

        private object? GetOptionListValueForCulture(string identifier, IEnumerable<IDataSourceValue> dataSourceValues, CultureInfo culture, CultureInfo? fallbackCulture)
        {
            // todo: deal with multi value
            var value = dataSourceValues.FirstOrDefault(x => x.Identifier == identifier);
            return value?.Labels.SingleOrDefault(x => x.Key == culture).Value ?? value?.Labels.SingleOrDefault(x => x.Key == fallbackCulture).Value ?? value?.Identifier;
        }

        private string GetEntityUrl(long entityId)
        {
            return string.Format(ApiConstants.ContentHubClient.EntityUrlFormat, entityId);
        }
    }
}
