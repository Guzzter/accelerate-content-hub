using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    public class PortalPage
    {
        public string Name { get; set; }

        public Dictionary<CultureInfo, string> Titles { get; set; } = [];
    }
}