using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;

namespace Sitecore.ContentHub.Integration.AssetImporter
{
    public class AssetImport(ILoggerFactory loggerFactory, IUploadWorker uploadWorker)
    {
        private readonly ILogger logger = loggerFactory.CreateLogger<AssetImport>();

        [Function("AssetImporterTimer")]
        public async Task RunTimer([TimerTrigger("0 */15 * * * *")] TimerInfo timer)
        {
            logger.LogInformation($"Asset Importer Timer trigger function executed at: {DateTime.Now}");

            try
            {
                await uploadWorker.ProcessUploads();
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing uploads");
            }

            if (timer.ScheduleStatus is not null)
                logger.LogInformation($"Next timer schedule at: {timer.ScheduleStatus.Next}");
        }

        [Function("AssetImporterHttp")]
        public async Task<IActionResult> RunHttp([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
        {
            try
            {
                await uploadWorker.ProcessUploads();
                return new OkObjectResult(new { success = true });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing uploads");
                return new BadRequestObjectResult(new { success = false });
            }
        }
    }
}
