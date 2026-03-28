using MiniExcelLibs.Attributes;
using System.ComponentModel;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    [DisplayName("M.Localization.Entry")]
    public class LocalizationEntry
    {
        [ExcelColumnName("id")]
        public required long Id { get; set; }

        [ExcelColumnName("identifier")]
        public required string Identifier { get; set; }

        [ExcelColumnName("M.Localization.Entry.Name")]
        public required string EntryName { get; set; }

        [ExcelColumnName("M.Localization.Entry.BaseTemplate")]
        public required string BaseTemplate { get; set; }

        [ExcelColumnName("M.Localization.Entry.Template")]
        public Dictionary<CultureInfo, string> Templates { get; set; } = [];
    }
}