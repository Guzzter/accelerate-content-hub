using Sitecore.ContentHub.Integration.SearchConnector.Models.Data;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface ISearchApiHelper
    {
        Task<string?> UpsertDocument(string entityId, DocumentData documentData);

        Task<IEnumerable<string>> UpsertDocuments(string entityId, IEnumerable<DocumentData> documentData);

        Task<IEnumerable<string>?> DeleteDocument(string entityId, string documentId, string locale = "all");
    }
}
