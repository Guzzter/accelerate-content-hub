using Stylelabs.M.Base.Querying;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    public interface IContentHubQueryBuilderService
    {
        IContentHubQueryBuilderService AddDefinitionFilter(string definitionName);

        IContentHubQueryBuilderService AddCreatedByFilter(long userId);

        IContentHubQueryBuilderService AddCreatedByFilter(string username);

        IContentHubQueryBuilderService AddPropertyFilter(string propertyName, string value);

        IContentHubQueryBuilderService AddPropertyFilter(string propertyName, bool value);

        IContentHubQueryBuilderService AddRelationFilter(string relationName, long id);

        IContentHubQueryBuilderService AddRelationFilter(string relationName, IEnumerable<long> ids);

        IContentHubQueryBuilderService Skip(int skip);

        IContentHubQueryBuilderService Take(int take);

        Query Build();
    }
}
