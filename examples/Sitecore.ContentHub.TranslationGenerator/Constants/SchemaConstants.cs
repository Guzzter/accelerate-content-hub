namespace Sitecore.CH.TranslationGenerator.Constants
{
    internal class SchemaConstants
    {
        internal class Shared
        {
            internal const char CultureSeparator = '#';

            internal class Properties
            {
                internal const string Id = "id";
                internal const string Identifier = "identifier";
            }
        }

        internal class LocalizationEntry
        {
            internal const string DefinitionName = "M.Localization.Entry";

            internal class Properties
            {
                internal const string Template = "M.Localization.Entry.Template";
                internal const string Name = "M.Localization.Entry.Name";
                internal const string BaseTemplate = "M.Localization.Entry.BaseTemplate";
            }
        }

        internal class PortalPage
        {
            internal const string DefinitionName = "Portal.Page";

            internal class Properties
            {
                internal const string Name = "Page.Name";
                internal const string Title = "Page.Title";
            }
        }

        internal class AssetType
        {
            internal const string DefinitionName = "M.AssetType";

            internal class Properties
            {
                internal const string Label = "Label";

                public static string Identifier { get; internal set; }
            }
        }
    }
}