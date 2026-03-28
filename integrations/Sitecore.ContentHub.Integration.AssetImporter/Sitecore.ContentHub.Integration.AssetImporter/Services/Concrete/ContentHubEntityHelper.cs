using Sitecore.ContentHub.Integration.AssetImporter.Constants;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Base.Querying.Linq;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ContentHubEntityHelper(IContentHubClientHelper clientHelper) : IContentHubEntityHelper
    {
        public async Task<long> GetIdFromIdentifierAsync(string identifier)
        {
            var entity = await clientHelper.Execute(client => client.Entities.GetAsync(identifier, EntityLoadConfiguration.Minimal));
            return entity == null || !entity.Id.HasValue
                ? throw new Exception($"Could not find entity with identifier {identifier}")
                : entity.Id.Value;
        }

        public async Task<long> GetIdFromUsernameAsync(string username)
        {
            var query = Query.CreateQuery(entities => (
                from e in entities
                where e.DefinitionName == SchemaConstants.User.DefinitionName && e.Property(SchemaConstants.User.Properties.Username) == username
                select e
            ));
            var result = await clientHelper.Execute(client => client.Querying.SingleIdAsync(query));
            return result == null ? throw new Exception($"Could not find user with username {username}") : result.Value;
        }
    }
}
