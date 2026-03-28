namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options
{
    public class OrderCloudOptions
    {
        public const string ConfigurationSectionName = "OrderCloud";

        public string? ApiUrl { get; set; }

        public string? AuthUrl { get; set; }

        public required string ClientId { get; set; }

        public required string ClientSecret { get; set; }
    }
}
