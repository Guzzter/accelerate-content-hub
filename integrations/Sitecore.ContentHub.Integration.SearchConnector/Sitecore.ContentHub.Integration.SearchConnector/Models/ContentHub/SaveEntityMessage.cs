namespace Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub
{
    public class SaveEntityMessage
    {
        public DateTime Timestamp { get; set; }

        public bool IsNew { get; set; }

        public required string TargetDefinition { get; set; }

        public long TargetId { get; set; }

        public required string TargetIdentifier { get; set; }

        public long UserId { get; set; }

        // ChangeSet
    }
}
