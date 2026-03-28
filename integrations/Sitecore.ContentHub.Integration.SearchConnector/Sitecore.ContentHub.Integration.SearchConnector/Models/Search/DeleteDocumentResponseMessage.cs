namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Search
{
    class DeleteDocumentResponseMessage
    {
        public bool Enqueued { get; set; }

        public IEnumerable<string>? IncrementalUpdateIds { get; set; }
    }
}