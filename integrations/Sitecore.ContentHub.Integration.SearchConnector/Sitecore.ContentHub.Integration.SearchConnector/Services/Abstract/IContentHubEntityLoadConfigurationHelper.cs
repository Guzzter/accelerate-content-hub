using Sitecore.ContentHub.Integration.SearchConnector.Models.Config;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;

namespace Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract
{
    interface IContentHubEntityLoadConfigurationHelper
    {
        IEntityLoadConfiguration BuildWithMappedCultures(IEnumerable<string> properties, IEnumerable<string> relations);

        IEntityLoadConfiguration BuildForFields(IEnumerable<FieldMap> fieldMaps);
    }
}
