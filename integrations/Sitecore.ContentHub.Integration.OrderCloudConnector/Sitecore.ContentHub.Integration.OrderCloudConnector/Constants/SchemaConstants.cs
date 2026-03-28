namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Constants
{
    class SchemaConstants
    {
        internal static class Catalog
        {
            internal const string DefinitionName = "M.PCM.Catalog";

            internal static class Properties
            {
                internal const string CatalogName = "CatalogName";
                internal const string CatalogDescription = "CatalogDescription";
            }

            internal static class Relations
            {
                internal const string CatalogToProduct = "PCMCatalogToProduct";
            }
        }

        internal static class Product
        {
            internal const string DefinitionName = "M.PCM.Product";

            internal static class Properties
            {
                internal const string ProductName = "ProductName";
                internal const string ProductLongDescription = "ProductLongDescription";
            }
        }
    }
}
