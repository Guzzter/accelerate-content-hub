using Azure.Messaging.ServiceBus;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options;
using System.Text.Json;
using FromBodyAttribute = Microsoft.Azure.Functions.Worker.Http.FromBodyAttribute;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector;

public class AddUpdateMessageFunction(ILogger<AddUpdateMessageFunction> logger, ServiceBusClient serviceBusClient, IOptions<ServiceBusOptions> serviceBusOptions)
{
    [Function(nameof(AddUpdateMessage))]
    public async Task<IActionResult> AddUpdateMessage([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] UpdateMessage updateMessage)
    {
        try
        {
            await serviceBusClient.CreateSender(serviceBusOptions.Value.UpdateQueueName).SendMessageAsync(new ServiceBusMessage(JsonSerializer.Serialize(updateMessage)));
            return new OkObjectResult(new { success = true });
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error adding update message");
            return new BadRequestObjectResult(new { success = false, error = ex.Message });
        }
    }
}