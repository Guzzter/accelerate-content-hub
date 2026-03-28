namespace Sitecore.ContentHub.Integration.AssetImporter.Models.Options
{
    public class AzureOptions
    {
        public const string ConfigurationSectionName = "Azure";

        public required string StorageContainerName { get; set; }
    }
}
