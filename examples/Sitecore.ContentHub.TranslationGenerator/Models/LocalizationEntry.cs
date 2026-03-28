using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    public class LocalizationEntry
    {
        public long Id { get; set; }
        public string Identifier { get; set; }
        public string EntryName { get; set; }
        public string BaseTemplate { get; set; }
        public Dictionary<CultureInfo, string> Templates { get; set; } = [];
    }
}