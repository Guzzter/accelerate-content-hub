using Azure.Messaging.ServiceBus;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using System.Text.Json;

namespace Sitecore.ContentHub.Integration.SearchConnector
{
    public class ServiceBusFunctions(ILogger<ServiceBusFunctions> logger, IMessageWorker messageWorker)
    {
        [Function(nameof(RunUpsert))]
        public async Task RunUpsert(
            [ServiceBusTrigger("%ServiceBus:UpsertQueueName%", Connection = "%ServiceBus:Name%")]
            ServiceBusReceivedMessage message,
            ServiceBusMessageActions messageActions)
        {
            try
            {
                var actionMessage = JsonSerializer.Deserialize<ActionMessage>(message.Body);
                if (actionMessage?.SaveEntityMessage == null)
                    throw new ArgumentException("Invalid request. SaveEntityMessage is null.");

                await messageWorker.UpsertSearchDocument(actionMessage.SaveEntityMessage.TargetDefinition, actionMessage.SaveEntityMessage.TargetIdentifier);
                await messageActions.CompleteMessageAsync(message);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing message: {MessageId}", message.MessageId);
                await messageActions.AbandonMessageAsync(message);
            }
        }

        [Function(nameof(RunDelete))]
        public async Task RunDelete(
            [ServiceBusTrigger("%ServiceBus:DeleteQueueName%", Connection = "%ServiceBus:Name%")]
            ServiceBusReceivedMessage message,
            ServiceBusMessageActions messageActions)
        {
            try
            {
                var actionMessage = JsonSerializer.Deserialize<ActionMessage>(message.Body);
                if (actionMessage?.SaveEntityMessage == null)
                    throw new ArgumentException("Invalid request. SaveEntityMessage is null.");

                await messageWorker.DeleteSearchDocument(actionMessage.SaveEntityMessage.TargetDefinition, actionMessage.SaveEntityMessage.TargetIdentifier);
                await messageActions.CompleteMessageAsync(message);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing message: {MessageId}", message.MessageId);
                await messageActions.AbandonMessageAsync(message);
            }
        }
    }
}
