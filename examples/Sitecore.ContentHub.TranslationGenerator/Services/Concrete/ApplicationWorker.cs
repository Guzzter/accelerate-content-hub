using Azure;
using Microsoft.Extensions.Hosting;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    public class ApplicationWorker(IConsoleHelper consoleHelper, IExcelService excelService, ITranslationService translationService) : IHostedService
    {
        private readonly IConsoleHelper consoleHelper = consoleHelper;
        private readonly IExcelService excelService = excelService;
        private readonly ITranslationService translationService = translationService;

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            consoleHelper.ResetStyle();
            consoleHelper.Write("PS - Content Hub - Translation Generator");

            while (true)
            {
                try
                {
                    var language = consoleHelper.GetInput("Target language", x => new CultureInfo(x), x => CultureInfo.GetCultures(CultureTypes.AllCultures).Any(c => c.Name == x.Name))!;
                    var inputFilePath = consoleHelper.GetInput("Input file", x => Path.GetExtension(x) == ".xlsx" && File.Exists(x));
                    var outputFilePath = consoleHelper.GetInput("Output file", GetDefaultOutputFilePath(inputFilePath, language.Name), x => Path.GetExtension(x) == ".xlsx" && !File.Exists(x));

                    var localisationEntries = excelService.GetLocalizationEntries(inputFilePath).ToList();
                    var portalPages = excelService.GetPortalPages(inputFilePath).ToList();
                    var assetTypes = excelService.GetAssetTypes(inputFilePath).ToList();

                    var translateLocalisationTasks = localisationEntries
                        .Where(x => !x.Templates.ContainsKey(language) || x.Templates[language] == string.Empty)
                        .Select(async x => x.Templates[language] = await translationService.Translate(language.Name, x.BaseTemplate));

                    var translatePagesTasks = portalPages
                        .Where(x => !x.Titles.ContainsKey(language) || x.Titles[language] == string.Empty)
                        .Select(async x => x.Titles[language] = await translationService.Translate(language.Name, x.Titles.First().Value, x.Titles.First().Key.Name));

                    var translateAssetTypesTasks = assetTypes
                        .Where(x => !x.Label.ContainsKey(language) || x.Label[language] == string.Empty)
                        .Select(async x => x.Label[language] = await translationService.Translate(language.Name, x.Label.First().Value, x.Label.First().Key.Name));

                    await Task.WhenAll(translateLocalisationTasks.Union(translatePagesTasks).Union(translateAssetTypesTasks));

                    excelService.Save(outputFilePath, localisationEntries, portalPages, assetTypes);
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

        private string GetDefaultOutputFilePath(string inputFilePath, string targetLanguage)
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