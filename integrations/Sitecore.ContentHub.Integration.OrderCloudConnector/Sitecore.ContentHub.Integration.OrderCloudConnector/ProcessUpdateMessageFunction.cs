using System.Text.Json;
using Azure.Messaging.ServiceBus;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector;

public class ProcessUpdateMessageFunction(ILogger<ProcessUpdateMessageFunction> logger, IMessageWorker messageWorker)
{
    [Function(nameof(RunUpdate))]
    public async Task RunUpdate(
        [ServiceBusTrigger("%ServiceBus:UpdateQueueName%", Connection = "%ServiceBus:Name%")]
        ServiceBusReceivedMessage message,
        ServiceBusMessageActions messageActions)
    {
        try
        {
            var updateMessage = JsonSerializer.Deserialize<UpdateMessage>(message.Body);
            if (updateMessage == null)
                throw new ArgumentException("Invalid message. Cannot deserialise to UpdateMessage.");

            await messageWorker.ProcessUpdateMessage(updateMessage);

            await messageActions.CompleteMessageAsync(message);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error processing message: {MessageId}", message.MessageId);
            await messageActions.AbandonMessageAsync(message);
        }
    }
}