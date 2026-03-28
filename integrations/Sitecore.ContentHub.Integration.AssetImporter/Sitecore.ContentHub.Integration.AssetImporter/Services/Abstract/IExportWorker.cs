using Sitecore.ContentHub.Integration.AssetImporter.Models.ContentHub;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    public interface IExportWorker
    {
        Task<IEnumerable<UploadedFileMetadata>> GenerateMetadataExport();

        Task<MemoryStream> GenerateMetadataExcel();
    }
}
