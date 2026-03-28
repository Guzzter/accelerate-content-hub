namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    interface IOrderCloudClientHelper
    {
        Task AssignProductToCatalog(string productId, string catalogId);

        Task SaveCatalog(string id, string name, string description, bool active = true);

        Task InactivateCatalog(string id);

        Task SaveProduct(string id, string name, string description, bool active = true);

        Task InactivateProduct(string id);
    }
}
