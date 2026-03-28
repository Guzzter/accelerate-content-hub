namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Data
{
    public class DocumentData
    {
        public required string Id { get; set; }

        public required string Locale { get; set; }

        public required IDictionary<string, object> Fields { get; set; }
    }
}
