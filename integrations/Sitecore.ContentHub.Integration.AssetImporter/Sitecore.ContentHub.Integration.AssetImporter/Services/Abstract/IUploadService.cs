namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    public interface IUploadService
    {
        Task<bool> Upload(string name, byte[] content);
    }
}
