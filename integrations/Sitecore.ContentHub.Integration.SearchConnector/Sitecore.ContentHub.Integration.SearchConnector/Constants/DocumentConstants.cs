using System.Text.RegularExpressions;

namespace Sitecore.ContentHub.Integration.SearchConnector.Constants
{
    internal static partial class DocumentConstants
    {
        [GeneratedRegex(@"\s+")]
        internal static partial Regex RepeatedWhitespaceRegex();

        internal const string ExtractedContentRenditionName = "downloadExtractedContent";

        internal const int ExtractedContentMaxLength = 20000;
    }
}
