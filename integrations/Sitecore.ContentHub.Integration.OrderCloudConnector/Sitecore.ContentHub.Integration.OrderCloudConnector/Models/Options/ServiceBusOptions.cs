namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options
{
    public class ServiceBusOptions
    {
        public const string ConfigurationSectionName = "ServiceBus";

        public required string Name { get; set; }

        public required string UpdateQueueName { get; set; }
    }
}
