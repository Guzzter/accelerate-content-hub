namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract
{
    public interface IExcelHelper
    {
        MemoryStream GenerateExcel<T>(IEnumerable<T> data, string? sheetName = null);
    }
}
