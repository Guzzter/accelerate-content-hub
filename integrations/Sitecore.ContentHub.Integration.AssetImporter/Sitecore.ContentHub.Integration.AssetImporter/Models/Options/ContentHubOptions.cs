namespace Sitecore.ContentHub.Integration.AssetImporter.Models.Options
{
    public class ContentHubOptions
    {
        public const string ConfigurationSectionName = "ContentHub";

        public required Uri BaseUrl { get; set; }
        public required string ClientId { get; set; }
        public required string ClientSecret { get; set; }
        public required string Username { get; set; }
        public required string Password { get; set; }
    }
}
