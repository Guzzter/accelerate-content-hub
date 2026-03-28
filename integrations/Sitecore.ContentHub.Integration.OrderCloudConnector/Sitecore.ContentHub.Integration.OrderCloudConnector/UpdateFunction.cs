using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;
using FromBodyAttribute = Microsoft.Azure.Functions.Worker.Http.FromBodyAttribute;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector;

public class UpdateFunction(ILogger<UpdateFunction> logger, IMessageWorker messageWorker)
{
    [Function(nameof(Update))]
    public async Task<IActionResult> Update([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] UpdateMessage updateMessage)
    {
        try
        {
            await messageWorker.ProcessUpdateMessage(updateMessage);

            return new OkObjectResult(new { success = true });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error processing http update message");
            return new BadRequestObjectResult(new { success = false, error = ex.Message });
        }
    }
}