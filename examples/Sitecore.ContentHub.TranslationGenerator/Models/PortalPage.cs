using MiniExcelLibs.Attributes;
using System.ComponentModel;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    [DisplayName("Portal.Page")]
    public class PortalPage
    {
        [ExcelColumnName("Page.Name")]
        public required string Name { get; set; }

        [ExcelColumnName("Page.Title")]
        public Dictionary<CultureInfo, string> Titles { get; set; } = [];
    }
}