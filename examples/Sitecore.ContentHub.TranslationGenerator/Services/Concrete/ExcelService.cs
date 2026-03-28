using Microsoft.Extensions.Logging;
using MiniExcelLibs;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    internal class ExcelService : IExcelService
    {
        private const char CultureSeparator = '#';
        private readonly ILogger<ExcelService> logger;

        public ExcelService(ILogger<ExcelService> logger)
        {
            this.logger = logger;
        }

        private static string GetSheetName<T>()
        {
            var attribute = typeof(T).GetCustomAttribute<DisplayNameAttribute>();
            return attribute?.DisplayName ?? typeof(T).Name;
        }

        private static string GetColumnName<T>(string propertyName)
        {
            var property = typeof(T).GetProperty(propertyName);
            var attribute = property?.GetCustomAttribute<ExcelColumnNameAttribute>();
            return attribute?.ExcelColumnName ?? propertyName;
        }

        public IEnumerable<LocalizationEntry>? GetLocalizationEntries(string filePath)
        {
            var sheetName = GetSheetName<LocalizationEntry>();
            try
            {
                using var stream = File.OpenRead(filePath);
                return stream.Query(true, sheetName).Select(x => GetLocalizationEntryFromExcelRow((IDictionary<string, object>)x)).ToList();
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Could not load rows for {SheetName}", sheetName);
            }
            return null;
        }

        public IEnumerable<PortalPage>? GetPortalPages(string filePath)
        {
            var sheetName = GetSheetName<PortalPage>();
            try
            {
                using var stream = File.OpenRead(filePath);
                return stream.Query(true, sheetName).Select(x => GetPortalPageFromExcelRow((IDictionary<string, object>)x)).ToList();
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Could not load rows for {SheetName}", sheetName);
            }
            return null;
        }

        public IEnumerable<AssetType>? GetAssetTypes(string filePath)
        {
            var sheetName = GetSheetName<AssetType>();
            try
            {
                using var stream = File.OpenRead(filePath);
                return stream.Query(true, sheetName).Select(x => GetAssetTypesFromExcelRow((IDictionary<string, object>)x)).ToList();
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Could not load rows for {SheetName}", sheetName);
            }
            return null;
        }

        public void Save(string filePath, IEnumerable<LocalizationEntry> localizationEntries, IEnumerable<PortalPage> portalPages, IEnumerable<AssetType> assetTypes)
        {
            using var stream = File.OpenWrite(filePath);
            var config = new OpenXmlConfiguration()
            {
                TableStyles = TableStyles.None
            };
            var data = new Dictionary<string, object>
            {
                { GetSheetName<LocalizationEntry>(), localizationEntries.Select(GetExcelRowFromLocalizationEntry) },
                { GetSheetName<PortalPage>(), portalPages.Select(GetExcelRowFromPortalPage) },
                { GetSheetName<AssetType>(), assetTypes.Select(GetExcelRowFromAssetType) }
            };
            stream.SaveAs(data, true, configuration: config);
        }

        private static LocalizationEntry GetLocalizationEntryFromExcelRow(IDictionary<string, object> row)
        {
            return new LocalizationEntry
            {
                Id = Convert.ToInt64((double)row[GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Id))]),
                Identifier = (string)row[GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Identifier))],
                EntryName = (string)row[GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.EntryName))],
                BaseTemplate = (string)row[GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.BaseTemplate))],
                Templates = row.Where(c => c.Key.StartsWith(GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Templates)))).ToDictionary(c => GetCultureInfoFromColumnName(c.Key), c => (string)c.Value)
            };
        }

        private static PortalPage GetPortalPageFromExcelRow(IDictionary<string, object> row)
        {
            return new PortalPage
            {
                Name = (string)row[GetColumnName<PortalPage>(nameof(PortalPage.Name))],
                Titles = row.Where(c => c.Key.StartsWith(GetColumnName<PortalPage>(nameof(PortalPage.Titles)))).ToDictionary(c => GetCultureInfoFromColumnName(c.Key), c => (string)c.Value)
            };
        }

        private AssetType GetAssetTypesFromExcelRow(IDictionary<string, object> row)
        {
            return new AssetType
            {
                Identifier = (string)row[GetColumnName<AssetType>(nameof(AssetType.Identifier))],
                Label = row.Where(c => c.Key.StartsWith(GetColumnName<AssetType>(nameof(AssetType.Label)))).ToDictionary(c => GetCultureInfoFromColumnName(c.Key), c => (string)c.Value)
            };
        }

        private IDictionary<string, object> GetExcelRowFromLocalizationEntry(LocalizationEntry localizationEntry)
        {
            var row = new Dictionary<string, object>()
            {
                { GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Id)), localizationEntry.Id },
                { GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Identifier)), localizationEntry.Identifier },
                { GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.EntryName)), localizationEntry.EntryName },
                { GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.BaseTemplate)), localizationEntry.BaseTemplate },
            };

            foreach (var template in localizationEntry.Templates)
                row.Add(GetColumnNameForCulture(GetColumnName<LocalizationEntry>(nameof(LocalizationEntry.Templates)), template.Key), template.Value);

            return row;
        }

        private IDictionary<string, object> GetExcelRowFromPortalPage(PortalPage portalPage)
        {
            var row = new Dictionary<string, object>()
            {
                { GetColumnName<PortalPage>(nameof(PortalPage.Name)), portalPage.Name },
            };

            foreach (var title in portalPage.Titles)
                row.Add(GetColumnNameForCulture(GetColumnName<PortalPage>(nameof(PortalPage.Titles)), title.Key), title.Value);

            return row;
        }

        private IDictionary<string, object> GetExcelRowFromAssetType(AssetType assetType)
        {
            var row = new Dictionary<string, object>()
            {
                { GetColumnName<AssetType>(nameof(AssetType.Identifier)), assetType.Identifier },
            };

            foreach (var field in assetType.Label)
                row.Add(GetColumnNameForCulture(GetColumnName<AssetType>(nameof(AssetType.Label)), field.Key), field.Value);

            return row;
        }

        private static CultureInfo GetCultureInfoFromColumnName(string columnName)
        {
            return new CultureInfo(columnName.Split(CultureSeparator).Last());
        }

        private static string GetColumnNameForCulture(string columnName, CultureInfo culture)
        {
            return $"{columnName}{CultureSeparator}{culture.Name}";
        }
    }
}