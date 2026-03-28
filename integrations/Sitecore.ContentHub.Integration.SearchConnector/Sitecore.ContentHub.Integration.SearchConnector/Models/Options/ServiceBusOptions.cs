namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Options
{
    public class ServiceBusOptions
    {
        public const string ConfigurationSectionName = "ServiceBus";

        public required string Name { get; set; }

        public required string DeleteQueueName { get; set; }

        public required string UpsertQueueName { get; set; }
    }
}
