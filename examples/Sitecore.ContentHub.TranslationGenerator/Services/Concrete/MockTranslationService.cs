using Sitecore.CH.TranslationGenerator.Constants;
using Sitecore.CH.TranslationGenerator.Services.Abstract;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    internal class MockTranslationService : ITranslationService
    {
        public Task<string> Translate(string targetLanguage, string text, string sourceLanguage = TranslationConstants.DefaultSourceLanguage)
        {
            return Task.FromResult( $"{targetLanguage}:{text}");
        }
    }
}
