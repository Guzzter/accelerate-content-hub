using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.SearchConnector.Constants;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Data;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Options;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Search;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using System.Net.Http.Json;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class SearchApiHelper(HttpClient httpClient, IOptions<SearchOptions> searchOptions) : ISearchApiHelper
    {
        public async Task<IEnumerable<string>?> DeleteDocument(string entityId, string documentId, string locale = ApiConstants.Search.AllLocale)
        {
            var response = await httpClient.DeleteAsync(GetDocumentUrl(entityId, documentId, locale));
            response.EnsureSuccessStatusCode();
            var responseMessage = await response.Content.ReadFromJsonAsync<DeleteDocumentResponseMessage>();
            return responseMessage?.IncrementalUpdateIds;
        }

        public async Task<string?> UpsertDocument(string entityId, DocumentData documentData)
        {
            var documentRequest = new DocumentRequestMessage
            {
                Document = new Document
                {
                    Id = documentData.Id,
                    Fields = documentData.Fields
                }
            };
            var response = await httpClient.PutAsJsonAsync(GetDocumentUrl(entityId, documentData.Id, documentData.Locale), documentRequest);
            response.EnsureSuccessStatusCode();
            var responseMessage = await response.Content.ReadFromJsonAsync<DocumentResponseMessage>();
            return responseMessage?.IncrementalUpdateId;
        }

        public async Task<IEnumerable<string>> UpsertDocuments(string entityId, IEnumerable<DocumentData> documentData)
        {
            var updateTasks = documentData.Select(dd => UpsertDocument(entityId, dd));
            var incrementalUpdateIds = await Task.WhenAll(updateTasks);
            return incrementalUpdateIds.Where(x => x != null).Cast<string>();
        }

        private string GetDocumentUrl(string entityId, string documentId, string locale)
        {
            return string.Format(ApiConstants.Search.DocumentUrlFormat, searchOptions.Value.DomainId, searchOptions.Value.SourceId, entityId, documentId, locale);
        }
    }
}
