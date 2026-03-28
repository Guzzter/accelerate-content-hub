using Sitecore.ContentHub.Integration.OrderCloudConnector.Constants;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options
{
    public class SchemaOptions
    {
        public const string ConfigurationSectionName = "Schema";

        public string CatalogEntityDefinition { get; set; } = SchemaConstants.Catalog.DefinitionName;

        public string CatalogNameProperty { get; set; } = SchemaConstants.Catalog.Properties.CatalogName;

        public string CatalogDescriptionProperty { get; set; } = SchemaConstants.Catalog.Properties.CatalogDescription;

        public string CatalogToProductRelation { get; set; } = SchemaConstants.Catalog.Relations.CatalogToProduct;

        public string ProductEntityDefinition { get; set; } = SchemaConstants.Product.DefinitionName;

        public string ProductNameProperty { get; set; } = SchemaConstants.Product.Properties.ProductName;

        public string ProductDescriptionProperty { get; set; } = SchemaConstants.Product.Properties.ProductLongDescription;
    }
}
