using Microsoft.Extensions.Logging;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using Sitecore.CH.TranslationGenerator.Services.Abstract;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    internal class ExcelService : IExcelService
    {
        private readonly ILogger<ExcelService> logger;

        public ExcelService(ILogger<ExcelService> logger)
        {
            this.logger = logger;
        }

        public Dictionary<string, object> GetAllSheets(string filePath)
        {
            var sheets = new Dictionary<string, object>();
            
            try
            {
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var sheetNames = stream.GetSheetNames();
                
                foreach (var sheetName in sheetNames)
                {
                    try
                    {
                        var rows = stream.Query(useHeaderRow: true, sheetName: sheetName)
                                         .Select(x => (IDictionary<string, object>)x)
                                         .ToList();
                        sheets.Add(sheetName, rows);
                    }
                    catch (Exception ex)
                    {
                        logger.LogWarning(ex, "Could not load rows for sheet '{SheetName}'", sheetName);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Could not open Excel file '{FilePath}'", filePath);
            }
            
            return sheets;
        }

        public void SaveAllSheets(string filePath, Dictionary<string, object> sheets)
        {
            using var stream = File.OpenWrite(filePath);
            var config = new OpenXmlConfiguration()
            {
                TableStyles = TableStyles.None
            };
            stream.SaveAs(sheets, configuration: config);
        }
    }
}