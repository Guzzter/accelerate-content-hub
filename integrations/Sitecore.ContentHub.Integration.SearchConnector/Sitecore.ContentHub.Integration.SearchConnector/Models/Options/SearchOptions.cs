namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Options
{
    public class SearchOptions
    {
        public const string ConfigurationSectionName = "Search";

        public required Uri BaseUrl { get; set; }
        public required string AuthToken { get; set; }
        public required long DomainId { get; set; }
        public required long SourceId { get; set; }
    }
}
