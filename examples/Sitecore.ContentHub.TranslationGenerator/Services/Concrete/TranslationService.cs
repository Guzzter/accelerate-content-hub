using System.Collections.Concurrent;
using Azure;
using Azure.AI.Translation.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    public class TranslationService : ITranslationService
    {
        private readonly ILogger<TranslationService> logger;
        private readonly TextTranslationClient client;
        private readonly string defaultSourceLanguage;
        private readonly ConcurrentDictionary<string, Lazy<Task<string>>> _cache = new();
        
        private long _totalCharactersSent;
        public long TotalCharactersSent => Interlocked.Read(ref _totalCharactersSent);

        public TranslationService(ILogger<TranslationService> logger, IOptionsMonitor<TranslationSettings> translationSettingsMonitor, IConfiguration configuration)
        {
            this.logger = logger;
            TranslationSettings settings = translationSettingsMonitor.CurrentValue;
            AzureKeyCredential credential = new(settings.ApiKey);
            client = new (credential, settings.Region);
            defaultSourceLanguage = configuration.GetValue<string>("SourceLanguage") ?? "en-US";
        }

        public Task<string> Translate(string targetLanguage, string text, string? sourceLanguage = null)
        {
            string source = string.IsNullOrEmpty(sourceLanguage) ? defaultSourceLanguage : sourceLanguage;
            string cacheKey = $"{targetLanguage}|{source}|{text}";

            var lazyTask = _cache.GetOrAdd(cacheKey, key => new Lazy<Task<string>>(() => TranslateInternal(targetLanguage, text, source)));

            return lazyTask.Value;
        }

        private async Task<string> TranslateInternal(string targetLanguage, string text, string source)
        {
            Interlocked.Add(ref _totalCharactersSent, text.Length);
            int maxRetries = 5;
            int delayMs = 1000;

            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    var response = await client.TranslateAsync(targetLanguage, text, source);
                    return response.Value[0].Translations[0].Text;
                }
                catch (RequestFailedException ex) when (ex.Status == 429)
                {
                    if (i == maxRetries - 1)
                    {
                        logger.LogError(ex, "Max retries reached after 429 errors. Translation failed for target language '{TargetLanguage}', text: '{Text}'", targetLanguage, text);
                        break;
                    }

                    logger.LogWarning("Rate limited (429) while translating to '{TargetLanguage}'. Retrying in {Delay}ms... (Attempt {Attempt}/{MaxAttempts})", targetLanguage, delayMs, i + 1, maxRetries);
                    await Task.Delay(delayMs);
                    delayMs *= 2; // Exponential backoff
                }
                catch (RequestFailedException ex)
                {
                    logger.LogWarning(ex, "Translation failed for target language '{TargetLanguage}', text: '{Text}'", targetLanguage, text);
                    break;
                }
            }
            
            // On failure, remove from cache so subsequent passes might try again
            string cacheKey = $"{targetLanguage}|{source}|{text}";
            _cache.TryRemove(cacheKey, out _);

            return string.Empty;
        }
    }
}
