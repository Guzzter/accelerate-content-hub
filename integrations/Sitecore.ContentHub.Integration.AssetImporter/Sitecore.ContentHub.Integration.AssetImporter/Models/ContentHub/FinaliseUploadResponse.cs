namespace Sitecore.ContentHub.Integration.AssetImporter.Models.ContentHub
{
    public class FinaliseUploadResponse
    {
        public bool Success { get; set; }
        public string? Message { get; set; }
        public long? AssetId { get; set; }
        public string? AssetIdentifier { get; set; }
    }
}
