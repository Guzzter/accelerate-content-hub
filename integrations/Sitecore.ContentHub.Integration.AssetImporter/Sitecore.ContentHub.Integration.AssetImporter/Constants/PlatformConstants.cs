namespace Sitecore.ContentHub.Integration.AssetImporter.Constants
{
    internal class PlatformConstants
    {
        internal class Client
        {
            internal const int RequestRetryAttempts = 5;
            internal const int RequestRetryDelayMin = 100;
            internal const int RequestRetryDelayMax = 2000;
        }

        internal class Upload
        {
            internal const long ChunkSize = 10000000;
        }
    }
}
