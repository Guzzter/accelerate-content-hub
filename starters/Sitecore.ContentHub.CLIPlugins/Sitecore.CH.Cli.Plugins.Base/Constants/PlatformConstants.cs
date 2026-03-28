namespace Sitecore.CH.Cli.Plugins.Base.Constants
{
    internal class PlatformConstants
    {
        internal class Client
        {
            internal const int RequestRetryAttempts = 5;
            internal const int RequestRetryDelayMin = 100;
            internal const int RequestRetryDelayMax = 2000;
        }

        internal class Query
        {
            internal const int MaxTake = 1000;
            internal const int MaxResultCount = 10000;
        }
    }
}
