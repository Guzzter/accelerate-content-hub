using Sitecore.CH.TranslationGenerator.Constants;

namespace Sitecore.CH.TranslationGenerator.Services.Abstract
{
    public interface ITranslationService
    {
        Task<string> Translate(string targetLanguage, string text, string sourceLanguage = TranslationConstants.DefaultSourceLanguage);
    }
}