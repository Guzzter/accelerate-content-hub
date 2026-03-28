using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Concrete
{
    class MessageWorker(ILogger<MessageWorker> logger, IOptions<SchemaOptions> schemaOptions, IContentHubEntityService contentHubEntityService, IOrderCloudClientHelper orderCloudClientHelper) : IMessageWorker
    {
        public async Task ProcessUpdateMessage(UpdateMessage updateMessage)
        {
            var updates = await PopulateEntityDefinitions(updateMessage);

            foreach (var update in updates)
            {
                logger.LogInformation($"Process update for entity {update.Identifier}, {update.EntityDefinition}, {update.Operation}");

                if (update.EntityDefinition == schemaOptions.Value.ProductEntityDefinition && update.Operation == EntityUpdateOperation.Update)
                    await UpdateProduct(update.Identifier);
                else if (update.EntityDefinition == schemaOptions.Value.ProductEntityDefinition && update.Operation == EntityUpdateOperation.Delete)
                    await orderCloudClientHelper.InactivateProduct(NormaliseIdentifier(update.Identifier));
                else if (update.EntityDefinition == schemaOptions.Value.CatalogEntityDefinition && update.Operation == EntityUpdateOperation.Update)
                    await UpdateCatalog(update.Identifier);
                else if (update.EntityDefinition == schemaOptions.Value.CatalogEntityDefinition && update.Operation == EntityUpdateOperation.Delete)
                    await orderCloudClientHelper.InactivateCatalog(NormaliseIdentifier(update.Identifier));
            }
        }

        private async Task<IEnumerable<EntityUpdate>> PopulateEntityDefinitions(UpdateMessage updateMessage)
        {
            var updatesToPopulate = updateMessage.Updates.Where(u => u.EntityDefinition == null);
            var identifiers = updatesToPopulate.Select(u => u.Identifier);
            var entities = await contentHubEntityService.GetEntityByIdentifier(identifiers, contentHubEntityService.BuildLoadConfiguration());
            entities.ToList().ForEach(e => updatesToPopulate.First(u => u.Identifier == e.Identifier).EntityDefinition = e.DefinitionName);
            return updateMessage.Updates;
        }

        private async Task UpdateProduct(string identifier)
        {
            logger.LogInformation($"Updating product with identifier: {identifier}");
            var entityLoadConfiguration = contentHubEntityService.BuildLoadConfiguration([schemaOptions.Value.ProductNameProperty, schemaOptions.Value.ProductDescriptionProperty], [schemaOptions.Value.CatalogToProductRelation]);
            var productEntity = await contentHubEntityService.GetEntityByIdentifier(identifier, entityLoadConfiguration);

            await orderCloudClientHelper.SaveProduct(
                NormaliseIdentifier(productEntity.Identifier),
                await contentHubEntityService.GetPropertyValue<string>(productEntity, schemaOptions.Value.ProductNameProperty),
                await contentHubEntityService.GetPropertyValue<string>(productEntity, schemaOptions.Value.ProductDescriptionProperty)
            );

            var relatedCatalogs = productEntity.GetRelation<IChildToManyParentsRelation>(schemaOptions.Value.CatalogToProductRelation).GetIds();
            await UpdateProductCatalogAssignment(identifier, relatedCatalogs);
        }

        private async Task UpdateProductCatalogAssignment(string productIdentifier, IEnumerable<long> catalogIds)
        {
            var catalogEntities = await contentHubEntityService.GetEntityById(catalogIds, contentHubEntityService.BuildLoadConfiguration());
            var assignmentTasks = catalogEntities.Select(c => orderCloudClientHelper.AssignProductToCatalog(NormaliseIdentifier(productIdentifier), c.Identifier));
            await Task.WhenAll(assignmentTasks);
        }

        private async Task UpdateCatalog(string identifier)
        {
            logger.LogInformation($"Updating catalog with identifier: {identifier}");
            var entityLoadConfiguration = contentHubEntityService.BuildLoadConfiguration([schemaOptions.Value.CatalogNameProperty, schemaOptions.Value.CatalogDescriptionProperty]);
            var catalogEntity = await contentHubEntityService.GetEntityByIdentifier(identifier, entityLoadConfiguration);

            await orderCloudClientHelper.SaveCatalog(
                NormaliseIdentifier(catalogEntity.Identifier),
                await contentHubEntityService.GetPropertyValue<string>(catalogEntity, schemaOptions.Value.CatalogNameProperty),
                await contentHubEntityService.GetPropertyValue<string>(catalogEntity, schemaOptions.Value.CatalogDescriptionProperty)
            );
        }

        private string NormaliseIdentifier(string identifier)
        {
            return identifier.Replace('.', '_');
        }
    }
}
