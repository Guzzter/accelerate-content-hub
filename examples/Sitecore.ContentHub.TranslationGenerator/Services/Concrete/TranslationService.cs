using Azure;
using Azure.AI.Translation.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.CH.TranslationGenerator.Constants;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    public class TranslationService : ITranslationService
    {
        private readonly ILogger<TranslationService> logger;
        private readonly TextTranslationClient client;

        public TranslationService(ILogger<TranslationService> logger, IOptionsMonitor<TranslationSettings> translationSettingsMonitor)
        {
            this.logger = logger;
            TranslationSettings settings = translationSettingsMonitor.CurrentValue;
            AzureKeyCredential credential = new(settings.ApiKey);
            client = new (credential, settings.Region);
        }

        public async Task<string> Translate(string targetLanguage, string text, string sourceLanguage = TranslationConstants.DefaultSourceLanguage)
        {
            try
            {
                var response = await client.TranslateAsync(targetLanguage, text, sourceLanguage);
                return response.Value[0].Translations[0].Text;
            }
            catch (RequestFailedException)
            {
                logger.LogWarning("A translation failed");
                return string.Empty;
            }
        }
    }
}
