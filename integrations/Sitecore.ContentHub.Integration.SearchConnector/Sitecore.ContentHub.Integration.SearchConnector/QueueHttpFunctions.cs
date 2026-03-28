using Azure.Messaging.ServiceBus;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Options;
using System.Text.Json;
using FromBodyAttribute = Microsoft.Azure.Functions.Worker.Http.FromBodyAttribute;

namespace Sitecore.ContentHub.Integration.SearchConnector
{
    public class QueueHttpFunctions(ILogger<QueueHttpFunctions> logger, ServiceBusClient serviceBusClient, IOptions<ServiceBusOptions> serviceBusOptions)
    {
        [Function("AddUpsertMessage")]
        public async Task<IActionResult> RunAdd([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] ActionMessage actionMessage)
        {
            try
            {
                await serviceBusClient.CreateSender(serviceBusOptions.Value.UpsertQueueName).SendMessageAsync(new ServiceBusMessage(JsonSerializer.Serialize(actionMessage)));
                return new OkObjectResult(new { success = true });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error adding upsert message");
                return new BadRequestObjectResult(new { success = false, error = ex.Message });
            }
        }

        [Function("AddDeleteMessage")]
        public async Task<IActionResult> RunDelete([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req, [FromBody] ActionMessage actionMessage)
        {
            try
            {
                await serviceBusClient.CreateSender(serviceBusOptions.Value.DeleteQueueName).SendMessageAsync(new ServiceBusMessage(JsonSerializer.Serialize(actionMessage)));
                return new OkObjectResult(new { success = true });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error adding delete message");
                return new BadRequestObjectResult(new { success = false, error = ex.Message });
            }
        }
    }
}
