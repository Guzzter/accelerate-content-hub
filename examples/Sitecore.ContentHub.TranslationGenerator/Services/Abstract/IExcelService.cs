using Sitecore.CH.TranslationGenerator.Models;

namespace Sitecore.CH.TranslationGenerator.Services.Abstract
{
    public interface IExcelService
    {
        Dictionary<string, object> GetAllSheets(string filePath);
        void SaveAllSheets(string filePath, Dictionary<string, object> sheets);
    }
}