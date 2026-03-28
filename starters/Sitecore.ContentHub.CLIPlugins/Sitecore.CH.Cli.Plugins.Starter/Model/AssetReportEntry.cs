namespace Sitecore.CH.Cli.Plugins.Starter.Model
{
    public class AssetReportEntry
    {
        public long EntityId { get; init; }
        public string Identifier { get; init; } = string.Empty;
        public string FileName { get; init; } = string.Empty;
        public string FinalLifecycle { get; init; } = string.Empty;
        public bool AssetMedia { get; set; }
        public bool FetchDone { get; set; }
        public bool RenditionsTriggered { get; set; }
        public bool PreviewRenditionDone { get; set; }
        public bool AllRenditionsCompleted { get; set; }
        public bool MetadataDone { get; set; }
    }
}
