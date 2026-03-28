namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IMessageWorker
    {
        public Task<IEnumerable<string>> UpsertSearchDocument(string contentHubEntityDefinition, string contentHubEntityIdentifier);

        public Task<IEnumerable<string>?> DeleteSearchDocument(string contentHubEntityDefinition, string contentHubEntityIdentifier);
    }
}
