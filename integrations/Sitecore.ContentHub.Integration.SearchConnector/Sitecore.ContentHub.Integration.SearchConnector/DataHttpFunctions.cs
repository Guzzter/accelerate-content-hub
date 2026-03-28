using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using FromBodyAttribute = Microsoft.Azure.Functions.Worker.Http.FromBodyAttribute;

namespace Sitecore.ContentHub.Integration.SearchConnector
{
    public class DataHttpFunctions(ILogger<DataHttpFunctions> logger, IMessageWorker messageWorker)
    {
        [Function("Upsert")]
        public async Task<IActionResult> RunUpsert([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] ActionMessage actionMessage)
        {
            try
            {
                if (actionMessage.SaveEntityMessage == null)
                    throw new ArgumentException("Invalid request. SaveEntityMessage is null.");

                await messageWorker.UpsertSearchDocument(actionMessage.SaveEntityMessage.TargetDefinition, actionMessage.SaveEntityMessage.TargetIdentifier);
                return new OkObjectResult(new { success = true });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing http message");
                return new BadRequestObjectResult(new { success = false, error = ex.Message });
            }
        }

        [Function("Delete")]
        public async Task<IActionResult> RunDelete([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] ActionMessage actionMessage)
        {
            try
            {
                if (actionMessage.SaveEntityMessage == null)
                    throw new ArgumentException("Invalid request. SaveEntityMessage is null.");

                await messageWorker.DeleteSearchDocument(actionMessage.SaveEntityMessage.TargetDefinition, actionMessage.SaveEntityMessage.TargetIdentifier);
                return new OkObjectResult(new { success = true });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing http message");
                return new BadRequestObjectResult(new { success = false, error = ex.Message });
            }
        }
    }
}
