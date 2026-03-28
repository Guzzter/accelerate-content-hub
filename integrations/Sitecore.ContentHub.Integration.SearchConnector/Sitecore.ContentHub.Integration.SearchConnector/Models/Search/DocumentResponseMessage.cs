namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Search
{
    class DocumentResponseMessage
    {
        public bool Enqueued { get; set; }

        public string? IncrementalUpdateId { get; set; }
    }
}
