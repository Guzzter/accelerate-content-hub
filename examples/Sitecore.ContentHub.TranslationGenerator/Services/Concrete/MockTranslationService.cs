using Sitecore.CH.TranslationGenerator.Services.Abstract;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    internal class MockTranslationService : ITranslationService
    {
        private long _totalCharactersSent;
        public long TotalCharactersSent => Interlocked.Read(ref _totalCharactersSent);

        public Task<string> Translate(string targetLanguage, string text, string? sourceLanguage = null)
        {
            Interlocked.Add(ref _totalCharactersSent, text.Length);
            return Task.FromResult( $"{targetLanguage}:{text}");
        }
    }
}
