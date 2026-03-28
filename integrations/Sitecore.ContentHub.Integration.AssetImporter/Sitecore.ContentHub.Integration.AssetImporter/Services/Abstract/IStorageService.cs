namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    public interface IStorageService
    {
        Task<IEnumerable<string>> GetFilesAsync();

        Task<byte[]> GetFileContentsAsync(string fileName);

        Task RemoveFileAsync(string fileName);
    }
}
