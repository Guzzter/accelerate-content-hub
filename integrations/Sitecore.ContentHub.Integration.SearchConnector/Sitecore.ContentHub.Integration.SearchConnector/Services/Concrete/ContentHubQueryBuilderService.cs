using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Base.Querying.Filters;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    internal class ContentHubQueryBuilderService : IContentHubQueryBuilderService
    {
        private Query query;

        public ContentHubQueryBuilderService()
        {
            query = new Query();
        }

        public IContentHubQueryBuilderService AddCreatedByFilter(long userId)
        {
            return AddFilter(new CreatedByQueryFilter
            {
                Operator = ComparisonOperator.Equals,
                Id = userId
            });
        }

        public IContentHubQueryBuilderService AddCreatedByFilter(string username)
        {
            return AddFilter(new CreatedByQueryFilter
            {
                Operator = ComparisonOperator.Equals,
                Username = username
            });
        }

        public IContentHubQueryBuilderService AddDefinitionFilter(string definitionName)
        {
            return AddFilter(new DefinitionQueryFilter
            {
                Operator = ComparisonOperator.Equals,
                Name = definitionName
            });
        }

        public IContentHubQueryBuilderService AddPropertyFilter(string propertyName, string value)
        {
            return AddPropertyFilter(propertyName, value, FilterDataType.String);
        }

        public IContentHubQueryBuilderService AddPropertyFilter(string propertyName, bool value)
        {
            return AddPropertyFilter(propertyName, value, FilterDataType.Bool);
        }

        public IContentHubQueryBuilderService AddRelationFilter(string relationName, long id)
        {
            return AddFilter(new RelationQueryFilter
            {
                Relation = relationName,
                ParentId = id
            });
        }

        public IContentHubQueryBuilderService AddRelationFilter(string relationName, IEnumerable<long> ids)
        {
            return AddFilter(new RelationQueryFilter
            {
                Relation = relationName,
                ParentIds = ids,
            });
        }

        public IContentHubQueryBuilderService Skip(int skip)
        {
            query.Skip = skip;
            return this;
        }

        public IContentHubQueryBuilderService Take(int take)
        {
            query.Take = take;
            return this;
        }

        public Query Build()
        {
            var returnQuery = query;
            query = new Query();
            return returnQuery;
        }

        private IContentHubQueryBuilderService AddPropertyFilter(string propertyName, object value, FilterDataType dataType)
        {
            return AddFilter(new PropertyQueryFilter
            {
                Operator = ComparisonOperator.Equals,
                Property = propertyName,
                DataType = dataType,
                Value = value
            });
        }

        private IContentHubQueryBuilderService AddFilter(QueryFilter filter)
        {
            if (query.Filter == null)
                query.Filter = filter;
            else if (query.Filter is CompositeQueryFilter compositeQueryFilter && compositeQueryFilter.CombineMethod == CompositeFilterOperator.And)
                compositeQueryFilter.Children.Add(filter);
            else
                query.Filter = new CompositeQueryFilter { Children = [query.Filter, filter], CombineMethod = CompositeFilterOperator.And };
            return this;
        }
    }
}
