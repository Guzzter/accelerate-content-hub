namespace Sitecore.ContentHub.Integration.AssetImporter.Constants
{
    internal class SchemaConstants
    {
        internal static class Asset
        {
            internal const string DefinitionName = "M.Asset";

            internal static class Properties
            {
                internal const string Filename = "Filename";
            }

            internal static class Relations
            {
                internal const string FinalLifeCycleStatusToAsset = "FinalLifeCycleStatusToAsset";
            }

            internal static class SystemProperties
            {
                internal const string CreatedOn = "created_on";
            }
        }

        internal static class User
        {
            internal const string DefinitionName = "User";

            internal static class Properties
            {
                internal const string Username = "username";
            }
        }
    }
}
