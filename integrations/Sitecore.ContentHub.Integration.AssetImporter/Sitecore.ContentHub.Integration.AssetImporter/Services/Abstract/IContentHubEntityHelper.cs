namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    public interface IContentHubEntityHelper
    {
        Task<long> GetIdFromIdentifierAsync(string identifier);

        Task<long> GetIdFromUsernameAsync(string username);
    }
}
