namespace Sitecore.CH.TranslationGenerator.Services.Abstract
{
    public interface ITranslationService
    {
        long TotalCharactersSent { get; }
        Task<string> Translate(string targetLanguage, string text, string? sourceLanguage = null);
    }
}