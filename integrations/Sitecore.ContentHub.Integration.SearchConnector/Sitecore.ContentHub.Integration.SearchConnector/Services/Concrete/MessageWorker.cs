using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;


namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class MessageWorker(ILogger<MessageWorker> logger, IContentHubEntityService entityService, IConfigHelper configHelper, ISearchApiHelper searchApiHelper, IContentHubEntityLoadConfigurationHelper entityLoadConfigurationHelper, IDocumentDataService documentDataService) : IMessageWorker
    {
        public async Task<IEnumerable<string>?> DeleteSearchDocument(string contentHubEntityDefinition, string contentHubEntityIdentifier)
        {
            logger.LogInformation($"Deleting search document for entity {contentHubEntityIdentifier}");
            var definitionMap = configHelper.GetDefinitionMap(contentHubEntityDefinition);
            return await searchApiHelper.DeleteDocument(definitionMap.SearchEntity, contentHubEntityIdentifier);
        }

        public async Task<IEnumerable<string>> UpsertSearchDocument(string contentHubEntityDefinition, string contentHubEntityIdentifier)
        {
            logger.LogInformation($"Upserting search document for entity {contentHubEntityIdentifier}");
            var definitionMap = configHelper.GetDefinitionMap(contentHubEntityDefinition);

            var loadConfiguration = entityLoadConfigurationHelper.BuildForFields(definitionMap.FieldMaps);
            var entity = await entityService.GetEntityByIdentifier(contentHubEntityIdentifier, loadConfiguration) ?? throw new InvalidDataException($"Entity with identifier {contentHubEntityIdentifier} not found.");
            
            var documentData = await documentDataService.GetDocumentData(definitionMap, entity);

            return await searchApiHelper.UpsertDocuments(definitionMap.SearchEntity, documentData);
        }
    }
}
