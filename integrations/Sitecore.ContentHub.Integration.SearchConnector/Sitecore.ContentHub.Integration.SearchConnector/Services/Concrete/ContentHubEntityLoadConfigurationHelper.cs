using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete
{
    class ContentHubEntityLoadConfigurationHelper(ICultureHelper cultureHelper) : IContentHubEntityLoadConfigurationHelper
    {
        public IEntityLoadConfiguration BuildWithMappedCultures(IEnumerable<string> properties, IEnumerable<string> relations)
        {
            return new EntityLoadConfigurationBuilder()
                .WithProperties(properties)
                .WithRelations(relations)
                .InCultures(cultureHelper.GetMappedContentHubCultures())
                .Build();
        }

        public IEntityLoadConfiguration BuildForFields(IEnumerable<FieldMap> fieldMaps)
        {
            var properties = fieldMaps
                .Where(x => x is PropertyFieldMap)
                .Cast<PropertyFieldMap>()
                .Select(x => x.ContentHubPropertyName);

            return new EntityLoadConfigurationBuilder()
                .WithProperties(properties)
                .InCultures(cultureHelper.GetMappedContentHubCultures())
                .Build();
        }
    }
}
