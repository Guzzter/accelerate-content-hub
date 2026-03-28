using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    public class AssetType
    {
        public string Identifier { get; set; }

        public Dictionary<CultureInfo, string> Label { get; set; } = [];
    }
}