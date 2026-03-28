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
            int maxRetries = 5;
            int delayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    var response = await client.TranslateAsync(targetLanguage, text, sourceLanguage);
                    return response.Value[0].Translations[0].Text;
                }
                catch (RequestFailedException ex) when (ex.Status == 429)
                {
                    if (i == maxRetries - 1)
                    {
                        logger.LogError(ex, "Max retries reached after 429 errors. Translation failed for target language '{TargetLanguage}', text: '{Text}'", targetLanguage, text);
                        return string.Empty;
                    }

                    logger.LogWarning("Rate limited (429) while translating to '{TargetLanguage}'. Retrying in {Delay}ms... (Attempt {Attempt}/{MaxAttempts})", targetLanguage, delayMs, i + 1, maxRetries);
                    await Task.Delay(delayMs);
                    delayMs *= 2; // Exponential backoff
                }
                catch (RequestFailedException ex)
                {
                    logger.LogWarning(ex, "Translation failed for target language '{TargetLanguage}', text: '{Text}'", targetLanguage, text);
                    return string.Empty;
                }
            }
            
            return string.Empty;
        }
    }
}
