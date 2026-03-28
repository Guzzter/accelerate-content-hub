namespace Sitecore.ContentHub.Integration.SearchConnector.Constants
{
    internal class ApiConstants
    {
        internal class ContentHubClient
        {
            internal const int RequestRetryAttempts = 5;
            internal const int RequestRetryDelayMin = 100;
            internal const int RequestRetryDelayMax = 2000;

            internal const int SearchAfterDefaultTake = 50;

            internal const string EntityUrlFormat = "/api/entities/{0}";

        }

        internal class Search
        {
            internal const string AllLocale = "all";
            internal const string DocumentUrlFormat = "ingestion/v1/domains/{0}/sources/{1}/entities/{2}/documents/{3}?locale={4}";
        }
    }
}
