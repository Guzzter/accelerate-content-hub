using MiniExcelLibs;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ExcelHelper : IExcelHelper
    {
        public MemoryStream GenerateExcel<T>(IEnumerable<T> data, string? sheetName = null)
        {
            var memoryStream = new MemoryStream();
            memoryStream.SaveAs(data, true, sheetName);
            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }
    }
}
