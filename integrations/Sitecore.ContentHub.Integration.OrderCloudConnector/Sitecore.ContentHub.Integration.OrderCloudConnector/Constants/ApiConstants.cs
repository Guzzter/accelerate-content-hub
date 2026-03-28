namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Constants
{
    internal class ApiConstants
    {
        internal class ContentHubClient
        {
            internal const int RequestRetryAttempts = 5;
            internal const int RequestRetryDelayMin = 100;
            internal const int RequestRetryDelayMax = 2000;

            internal const int SearchAfterDefaultTake = 50;
        }
    }
}
