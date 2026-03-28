namespace Sitecore.ContentHub.Integration.AssetImporter.Constants
{
    internal class ApiConstants
    {
        internal class Endpoints
        {
            internal const string UploadRequest = "/api/v2.0/upload";
            internal const string UploadFinalise = "/api/v2.0/upload/finalize";
        }

        internal class UploadConfiguration
        {
            internal const string AssetUpload = "AssetUploadConfiguration";
        }

        internal class UploadAction
        {
            internal const string NewAsset = "NewAsset";
        }

        internal class FormDataKeys
        {
            internal const string UploadFile = "file";
        }
    }
}
