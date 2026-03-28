using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Data;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IDocumentDataService
    {
        Task<IEnumerable<DocumentData>> GetDocumentData(DefinitionMap definitionMap, IEntity entity);
    }
}
