using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;

namespace Sitecore.ContentHub.Integration.AssetImporter
{
    public class MetadataExport(ILoggerFactory loggerFactory, IExportWorker exportWorker)
    {
        private readonly ILogger logger = loggerFactory.CreateLogger<MetadataExport>();

        [Function("MetadataExport")]
        public async Task<IActionResult> Export([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
        {
            try
            {
                return new OkObjectResult(await exportWorker.GenerateMetadataExport());
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error generating metadata");
                return new BadRequestObjectResult(new { success = false });
            }
        }

        [Function("MetadataExportExcel")]
        public async Task<IActionResult> ExportExcel([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
        {
            try
            {
                return new FileStreamResult(await exportWorker.GenerateMetadataExcel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {  FileDownloadName = "Export.xlsx" };
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error generating metadata excel");
                return new BadRequestObjectResult(new { success = false });
            }
        }
    }
}
