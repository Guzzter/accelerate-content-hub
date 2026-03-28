namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Search
{
    public class Document
    {
        public required string Id { get; set; }
        public required IDictionary<string, object> Fields { get; set; } = new Dictionary<string, object>();
    }
}
