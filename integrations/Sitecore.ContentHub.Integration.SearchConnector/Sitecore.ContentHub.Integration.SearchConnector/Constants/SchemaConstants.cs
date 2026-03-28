namespace Sitecore.ContentHub.Integration.SearchConnector.Constants
{
    class SchemaConstants
    {
        internal static class SystemProperties
        {
            internal const string Identifier = "identifier";
        }

        internal static class Asset
        {
            internal const string DefinitionName = "M.Asset";

            internal static class Relations
            {
                internal const string AssetToPublicLink = "AssetToPublicLink";
            }
        }

        internal static class PublicLink
        {
            internal const string DefinitionName = "M.PublicLink";

            internal static class Properties
            {
                internal const string RelativeUrl = "RelativeUrl";
                internal const string Resource = "Resource";
                internal const string VersionHash = "VersionHash";
                internal const string IsDisabled = "IsDisabled";
            }

            internal static class Relations
            {
                internal const string AssetToPublicLink = "AssetToPublicLink";
            }
        }
    }
}
