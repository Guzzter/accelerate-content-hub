using Microsoft.Extensions.Logging;
using OrderCloud.SDK;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Concrete
{
    class OrderCloudClientHelper(ILogger<OrderCloudClientHelper> logger, IOrderCloudClientFactory clientFactory) : IOrderCloudClientHelper
    {
        private readonly OrderCloudClient client = clientFactory.CreateClient();

        public async Task AssignProductToCatalog(string productId, string catalogId)
        {
            try
            {
                logger.LogInformation($"Assigning product {productId} to catalog {catalogId}");
                var assignment = new ProductCatalogAssignment
                {
                    ProductID = productId,
                    CatalogID = catalogId
                };
                await client.Catalogs.SaveProductAssignmentAsync(assignment);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Failed to assign product {productId} to catalog {catalogId}");
            }
        }

        public async Task SaveCatalog(string id, string name, string description, bool active = true)
        {
            try
            {
                logger.LogInformation($"Saving catalog {id}");
                var catalog = new Catalog
                {
                    ID = id,
                    Name = name,
                    Description = description,
                    Active = active,
                };
                await client.Catalogs.SaveAsync(id, catalog);
            }
            catch(Exception ex)
            {
                logger.LogError(ex, $"Failed to save catalog {id}");
            }
        }

        public async Task InactivateCatalog(string id)
        {
            try
            {
                logger.LogInformation($"Inactivating catalog {id}");
                await client.Catalogs.PatchAsync(id, new PartialCatalog
                {
                    Active = false
                });
            }
            catch(OrderCloudException ex)
            {
                if (ex.HttpStatus == System.Net.HttpStatusCode.NotFound)
                    logger.LogWarning($"Failed to inactivate catalog {id}: NotFound");
                else
                    logger.LogError(ex, $"Failed to inactivate catalog {id}"); ;
            }
        }

        public async Task SaveProduct(string id, string name, string description, bool active = true)
        {
            try
            {
                logger.LogInformation($"Saving product {id}");
                var product = new Product
                {
                    ID = id,
                    Name = name,
                    Description = description,
                    Active = active,
                };

                await client.Products.SaveAsync(id, product);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Failed to save product {id}");
            }
        }

        public async Task InactivateProduct(string id)
        {
            try
            {
                logger.LogInformation(id, $"Inactivating product {id}");
                await client.Products.PatchAsync(id, new PartialProduct
                {
                    Active = false
                });
            }
            catch (OrderCloudException ex)
            {
                if (ex.HttpStatus == System.Net.HttpStatusCode.NotFound)
                    logger.LogWarning($"Failed to inactivate product {id}: NotFound");
                else
                    logger.LogError(ex, $"Failed to inactivate product {id}"); ;
            }
        }
    }
}
