using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    public class ApplicationWorker(IConsoleHelper consoleHelper, IExcelService excelService, ITranslationService translationService, Microsoft.Extensions.Configuration.IConfiguration configuration) : IHostedService
    {
        private const string ExcelExtension = ".xlsx";
        private const int MaxConcurrency = 10;
        private static readonly SemaphoreSlim _throttle = new(MaxConcurrency);

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            consoleHelper.ResetStyle();
            consoleHelper.Write("PS - Content Hub - Translation Generator");

            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    var targetLanguages = configuration.GetSection("TargetLanguages").Get<string[]>() ?? [];
                    List<CultureInfo> selectedLanguages = [];

                    if (targetLanguages.Length > 0)
                    {
                        var dictionary = targetLanguages.Select((lang, i) => new { Index = i + 1, Lang = lang }).ToDictionary(x => x.Index, x => x.Lang);
                        consoleHelper.Write("Available languages:");
                        foreach (var kvp in dictionary)
                            consoleHelper.Write($"{kvp.Key} - {kvp.Value}");
                        consoleHelper.Write("A - All");

                        var selection = consoleHelper.GetInput("Select languages (comma separated, e.g. 1, 2) or 'A' for all", x => x, x => {
                            if (x.Trim().Equals("A", StringComparison.OrdinalIgnoreCase)) return true;
                            var parts = x.Split(',');
                            return parts.All(p => int.TryParse(p.Trim(), out var pInt) && dictionary.ContainsKey(pInt));
                        })!;

                        if (selection.Trim().Equals("A", StringComparison.OrdinalIgnoreCase))
                            selectedLanguages = targetLanguages.Select(x => new CultureInfo(x)).ToList();
                        else
                        {
                            var indices = selection.Split(',').Select(x => int.Parse(x.Trim()));
                            selectedLanguages = indices.Select(i => new CultureInfo(dictionary[i])).ToList();
                        }
                    }
                    else
                    {
                        var languageStr = consoleHelper.GetInput("Target languages (comma separated, e.g. de-DE, nl-NL)", x => x, x => x.Split(',').All(p => CultureInfo.GetCultures(CultureTypes.AllCultures).Any(c => c.Name == p.Trim())))!;
                        selectedLanguages = languageStr.Split(',').Select(x => new CultureInfo(x.Trim())).ToList();
                    }

                    var inputFilePath = consoleHelper.GetInput("Input file", x => Path.GetExtension(x) == ExcelExtension && File.Exists(x))!;
                    
                    string languageLabel = selectedLanguages.Count == 1 ? selectedLanguages[0].Name : "Multiple";
                    var outputFilePath = consoleHelper.GetInput("Output file", GetDefaultOutputFilePath(inputFilePath, languageLabel), x => Path.GetExtension(x) == ExcelExtension && !File.Exists(x))!;

                    var localisationEntries = excelService.GetLocalizationEntries(inputFilePath)?.ToList();
                    var portalPages = excelService.GetPortalPages(inputFilePath)?.ToList();
                    var assetTypes = excelService.GetAssetTypes(inputFilePath)?.ToList();

                    if (localisationEntries == null || portalPages == null || assetTypes == null)
                    {
                        consoleHelper.Write("Failed to load one or more data sheets from the input file. Please check the file format.");
                        continue;
                    }

                    foreach (var language in selectedLanguages)
                    {
                        await TranslateAllAsync(language, localisationEntries, portalPages, assetTypes, cancellationToken);
                    }

                    excelService.Save(outputFilePath, localisationEntries, portalPages, assetTypes);
                    consoleHelper.Write($"Output saved to: {outputFilePath}");
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    consoleHelper.Write("An error occurred:");
                    consoleHelper.Write(ex.ToString());
                }

                if (consoleHelper.GetExpectedChar("Press X to exit, or any other key to continue", 'x'))
                    break;
            }
        }

        private async Task TranslateAllAsync(
            CultureInfo language,
            List<LocalizationEntry> localisationEntries,
            List<PortalPage> portalPages,
            List<AssetType> assetTypes,
            CancellationToken cancellationToken)
        {
            var tasks = new List<Task>();

            List<LocalizationEntry> localisationToTranslate = [];/* = localisationEntries
                .Where(x => !x.Templates.ContainsKey(language) || x.Templates[language] == string.Empty)
                .ToList();*/

            List<PortalPage> pagesToTranslate = []; /* = portalPages
                .Where(x => !x.Titles.ContainsKey(language) || x.Titles[language] == string.Empty)
                .Where(x => x.Titles.Count > 0)
                .ToList();*/

            List<AssetType> assetTypesToTranslate = assetTypes
                .Where(x => !x.Label.ContainsKey(language) || x.Label[language] == string.Empty)
                .Where(x => x.Label.Count > 0)
                .ToList();

            int totalItems = localisationToTranslate.Count + pagesToTranslate.Count + assetTypesToTranslate.Count;
            int[] progress = [0];

            if (totalItems == 0)
            {
                consoleHelper.Write("All items are already translated.");
                return;
            }

            consoleHelper.Write($"Translating {totalItems} items to {language.Name}...");

            tasks.AddRange(localisationToTranslate.Select(x => ThrottledTranslateAsync(async () =>
            {
                x.Templates[language] = await translationService.Translate(language.Name, x.BaseTemplate);
                consoleHelper.OverwriteLine($"Progress: {Interlocked.Increment(ref progress[0])}/{totalItems}");
            }, cancellationToken)));

            tasks.AddRange(pagesToTranslate.Select(x => ThrottledTranslateAsync(async () =>
            {
                var source = x.Titles.First();
                x.Titles[language] = await translationService.Translate(language.Name, source.Value, source.Key.Name);
                consoleHelper.OverwriteLine($"Progress: {Interlocked.Increment(ref progress[0])}/{totalItems}");
            }, cancellationToken)));

            tasks.AddRange(assetTypesToTranslate.Select(x => ThrottledTranslateAsync(async () =>
            {
                var source = x.Label.First();
                x.Label[language] = await translationService.Translate(language.Name, source.Value, source.Key.Name);
                consoleHelper.OverwriteLine($"Progress: {Interlocked.Increment(ref progress[0])}/{totalItems}");
            }, cancellationToken)));

            await Task.WhenAll(tasks);
            consoleHelper.Write($"\nTranslation complete. {progress[0]}/{totalItems} items translated.");
        }

        private static async Task ThrottledTranslateAsync(Func<Task> work, CancellationToken cancellationToken)
        {
            await _throttle.WaitAsync(cancellationToken);
            try
            {
                await work();
            }
            finally
            {
                _throttle.Release();
            }
        }

        private static string GetDefaultOutputFilePath(string inputFilePath, string targetLanguage)
        {
            var outputFilePath = $"{Path.GetDirectoryName(inputFilePath)}{Path.DirectorySeparatorChar}{Path.GetFileNameWithoutExtension(inputFilePath)} ({targetLanguage}){Path.GetExtension(inputFilePath)}";
            if (File.Exists(outputFilePath))
            {
                int i = 1;
                do
                {
                    outputFilePath = $"{Path.GetDirectoryName(inputFilePath)}{Path.DirectorySeparatorChar}{Path.GetFileNameWithoutExtension(inputFilePath)} ({targetLanguage} - {i}){Path.GetExtension(inputFilePath)}";
                    i++;
                }
                while (File.Exists(outputFilePath));
            }
            return outputFilePath;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}