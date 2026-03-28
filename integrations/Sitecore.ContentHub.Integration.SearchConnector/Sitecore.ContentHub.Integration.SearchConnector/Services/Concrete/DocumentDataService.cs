using Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Data;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using System.Globalization;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Stylelabs.M.Sdk.Models.Typed;
using System.Text.RegularExpressions;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    partial class DocumentDataService(ILogger<DocumentDataService> logger, ICultureHelper cultureHelper, IContentHubClientHelper clientHelper, IContentHubEntityService entityService, IContentHubAssetService assetService) : IDocumentDataService
    {
        public async Task<IEnumerable<DocumentData>> GetDocumentData(DefinitionMap definitionMap, IEntity entity)
        {
            logger.LogInformation($"Getting data for entity {entity.Identifier}");
            var cultureMaps = cultureHelper.GetCultureMaps(entity.Cultures);
            var defaultCulture = await cultureHelper.GetContentHubDefaultCulture();
            var documentDataDictionary = cultureMaps.ToDictionary(x => x.ContentHubCulture, x => new DocumentData { Id = entity.Identifier, Locale = x.SearchCulture, Fields = new Dictionary<string, object>() });
            var documentContentDictionary = new Dictionary<string, object>();

            foreach (var fieldMap in definitionMap.FieldMaps)
            {
                if (fieldMap is PropertyFieldMap propertyFieldMap)
                    AddPropertyFieldData(documentDataDictionary, entity, propertyFieldMap, defaultCulture);

                if (fieldMap is OptionListFieldMap optionListFieldMap)
                    await AddOptionListFieldData(documentDataDictionary, entity, optionListFieldMap, defaultCulture);

                else if (fieldMap is RelationFieldMap relationFieldMap)
                    await AddRelationFieldData(documentDataDictionary, entity, relationFieldMap, defaultCulture);

                else if (fieldMap is PublicLinkFieldMap publicLinkFieldMap)
                    await AddPublicLinkFieldData(documentDataDictionary, entity, publicLinkFieldMap);

                else if (fieldMap is ExtractedContentFieldMap extractedContentFieldMap)
                    await AddExtractedContentFieldData(documentDataDictionary, entity, extractedContentFieldMap);
            }

            return documentDataDictionary.Values;
        }

        private void AddPropertyFieldData(Dictionary<string, DocumentData> documentDataDictionary, IEntity entity, PropertyFieldMap propertyFieldMap, CultureInfo defaultCulture)
        {
            var propertyValues = entityService
                .GetPropertyValueForCultures(entity, propertyFieldMap.ContentHubPropertyName, defaultCulture)
                .Where(x => x.Value != null)
                .Cast<KeyValuePair<string, object>>();

            foreach (var propertyValue in propertyValues)
                AddToDocumentData(documentDataDictionary, propertyValue.Key, propertyFieldMap.SearchAttributeName, propertyValue.Value);
        }

        private async Task AddOptionListFieldData(Dictionary<string, DocumentData> documentDataDictionary, IEntity entity, OptionListFieldMap optionListFieldMap, CultureInfo defaultCulture)
        {
            var optionList = await clientHelper.Execute(client => client.DataSources.GetAsync(optionListFieldMap.ContentHubOptionListName));

            var propertyValues = entityService
                .GetOptionListValueForCultures(entity, optionListFieldMap.ContentHubPropertyName, optionList, defaultCulture)
                .Where(x => x.Value != null)
                .Cast<KeyValuePair<string, object>>();

            foreach (var propertyValue in propertyValues)
                AddToDocumentData(documentDataDictionary, propertyValue.Key, optionListFieldMap.SearchAttributeName, propertyValue.Value);
        }

        private async Task AddRelationFieldData(Dictionary<string, DocumentData> documentDataDictionary, IEntity entity, RelationFieldMap relationFieldMap, CultureInfo defaultCulture)
        {
            var entityLoadConfiguration = new EntityLoadConfigurationBuilder()
                .WithProperty(relationFieldMap.ContentHubRelatedPropertyName)
                .InCultures(cultureHelper.GetMappedContentHubCultures())
                .Build();

            var relatedEntities = await entityService.GetRelatedEntities(entity, relationFieldMap.ContentHubRelations, entityLoadConfiguration);

            var relatedPropertyValues = relatedEntities
                .SelectMany(x => entityService.GetPropertyValueForCultures(x, relationFieldMap.ContentHubRelatedPropertyName, defaultCulture))
                .Where(x => x.Value != null)
                .GroupBy(x => x.Key, x => x.Value)
                .ToList();

            foreach (var group in relatedPropertyValues)
                AddToDocumentData(documentDataDictionary, group.Key, relationFieldMap.SearchAttributeName, group.ToList());
        }

        private async Task AddPublicLinkFieldData(Dictionary<string, DocumentData> documentDataDictionary, IEntity entity, PublicLinkFieldMap publicLinkFieldMap)
        {
            var assetEntity = await GetRelatedAsset(entity, publicLinkFieldMap);

            var publicLink = await assetService.GetPublicLink(assetEntity, publicLinkFieldMap.ContentHubResourceName, publicLinkFieldMap.CreateLinkIfNotExists);

            if (publicLink != null)
                foreach (var culture in documentDataDictionary.Keys)
                    AddToDocumentData(documentDataDictionary, culture, publicLinkFieldMap.SearchAttributeName, publicLink.ToString());
        }

        private async Task AddExtractedContentFieldData(Dictionary<string, DocumentData> documentDataDictionary, IEntity entity, ExtractedContentFieldMap extractedContentFieldMap)
        {
            var assetEntity = await GetRelatedAsset(entity, extractedContentFieldMap);

            var extractedContent = await assetService.GetExtractedContent(assetEntity);
            if (extractedContent != null)
            {
                var noWhitespaceContent = DocumentConstants.RepeatedWhitespaceRegex().Replace(extractedContent, " ").Trim();
                var trimmedContent = noWhitespaceContent[..Math.Min(noWhitespaceContent.Length, DocumentConstants.ExtractedContentMaxLength)];
                foreach (var culture in documentDataDictionary.Keys)
                    AddToDocumentData(documentDataDictionary, culture, extractedContentFieldMap.SearchAttributeName, trimmedContent);
            }
        }

        private async Task<IEntity> GetRelatedAsset(IEntity entity, OptionalRelationBaseFieldMap relationFieldMap)
        {
            IEntity assetEntity;
            if (relationFieldMap.ContentHubRelations != null)
            {
                var relatedEntities = await entityService.GetRelatedEntities(entity, relationFieldMap.ContentHubRelations);
                if (relatedEntities == null || relatedEntities.Count() != 1)
                    throw new Exception($"Relation field {string.Join('>', relationFieldMap.ContentHubRelations.Select(x => x.Name))} must have exactly one related entity.");
                assetEntity = relatedEntities.First();
            }
            else
                assetEntity = entity;

            if (assetEntity == null || !assetEntity.Id.HasValue || assetEntity.DefinitionName != SchemaConstants.Asset.DefinitionName)
                throw new Exception("Relation field must be related to an asset entity.");
            return assetEntity;
        }

        private static void AddToDocumentData(Dictionary<string, DocumentData> dict, string cultureKey, string fieldName, object value)
        {
            if (dict.TryGetValue(cultureKey, out var doc))
                doc.Fields[fieldName] = value;
        }
    }
}
