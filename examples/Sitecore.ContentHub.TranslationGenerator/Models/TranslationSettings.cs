namespace Sitecore.CH.TranslationGenerator.Models
{
    public class TranslationSettings
    {
        public const string Key = "Translation";

        public required string ApiKey { get; set; }
        public required string Region { get; set; }
    }
}
