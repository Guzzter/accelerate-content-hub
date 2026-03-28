using MiniExcelLibs.Attributes;
using System.ComponentModel;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Models
{
    [DisplayName("M.AssetType")]
    public class AssetType
    {
        [ExcelColumnName("Identifier")]
        public required string Identifier { get; set; }

        [ExcelColumnName("Label")]
        public Dictionary<CultureInfo, string> Label { get; set; } = [];
    }
}