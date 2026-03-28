using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;
using System.Globalization;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    public class ApplicationWorker(IConsoleHelper consoleHelper, IExcelService excelService, ITranslationService translationService, IConfiguration configuration) : IHostedService
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
                    var sourceLanguageStr = configuration.GetValue<string>("SourceLanguage") ?? "en-US";
                    var sourceLanguage = new CultureInfo(sourceLanguageStr);

                    var targetLanguages = (configuration.GetSection("TargetLanguages").Get<string[]>() ?? []).OrderBy(x => x).ToArray();
                    var definitionsToTranslate = configuration.GetSection("DefinitionsToTranslate").Get<string[]>() ?? [];
                    var definitionsToCopyValues = configuration.GetSection("DefinitionsToCopyValues").Get<string[]>() ?? [];

                    List<CultureInfo> selectedLanguages = [];

                    if (targetLanguages.Length > 0)
                    {
                        var dictionary = targetLanguages.Select((lang, i) => new { Index = i + 1, Lang = lang }).ToDictionary(x => x.Index, x => x.Lang);
                        consoleHelper.Write("Available languages:");
                        foreach (var kvp in dictionary)
                            consoleHelper.Write($"{kvp.Key} - {kvp.Value}");
                        consoleHelper.Write("A - All");

                        var selection = consoleHelper.GetInput("Select languages (comma separated, e.g. 1, 2) or 'A' for all", x => x, x =>
                        {
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

                    var sheets = excelService.GetAllSheets(inputFilePath);

                    if (sheets == null || sheets.Count == 0)
                    {
                        consoleHelper.Write("Failed to load any sheets from the input file. Please check the file format.");
                        continue;
                    }

                    foreach (var language in selectedLanguages)
                    {
                        await TranslateAllDynamicAsync(language, sourceLanguage, sheets, definitionsToTranslate, definitionsToCopyValues, cancellationToken);
                    }

                    var filteredSheets = sheets
                        .Where(kvp => definitionsToTranslate.Contains(kvp.Key, StringComparer.OrdinalIgnoreCase) || 
                                     definitionsToCopyValues.Contains(kvp.Key, StringComparer.OrdinalIgnoreCase))
                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                    excelService.SaveAllSheets(outputFilePath, filteredSheets);
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

        private async Task TranslateAllDynamicAsync(
            CultureInfo targetLanguage,
            CultureInfo sourceLanguage,
            Dictionary<string, object> sheets,
            string[] definitionsToTranslate,
            string[] definitionsToCopyValues,
            CancellationToken cancellationToken)
        {
            var tasks = new List<Task>();
            int totalItems = 0;
            int copiedItems = 0;
            int[] progress = [0];

            List<Func<Task>> translationActions = [];

            foreach (var kvp in sheets)
            {
                string sheetName = kvp.Key;

                bool shouldTranslate = definitionsToTranslate.Contains(sheetName, StringComparer.OrdinalIgnoreCase);
                bool shouldCopy = definitionsToCopyValues.Contains(sheetName, StringComparer.OrdinalIgnoreCase);

                if (!shouldTranslate && !shouldCopy)
                    continue;

                if (kvp.Value is List<IDictionary<string, object>> rows)
                {
                    string sourceSuffix = $"#{sourceLanguage.Name}";
                    string targetSuffix = $"#{targetLanguage.Name}";

                    // Identify all possible target columns for this sheet and initialize them for all rows
                    // This ensures MiniExcel has a consistent schema across all rows, preventing KeyNotFoundException.
                    var targetColumns = rows.SelectMany(r => r.Keys)
                                            .Where(k => k.EndsWith(sourceSuffix, StringComparison.OrdinalIgnoreCase))
                                            .Select(k => $"{k.Substring(0, k.Length - sourceSuffix.Length)}{targetSuffix}")
                                            .Distinct()
                                            .ToList();

                    foreach (var row in rows)
                    {
                        foreach (var col in targetColumns)
                        {
                            if (!row.ContainsKey(col)) row[col] = null;
                        }
                    }

                    foreach (var row in rows)
                    {
                        var sourceKeys = row.Keys.Where(k => k.EndsWith(sourceSuffix, StringComparison.OrdinalIgnoreCase)).ToList();

                        foreach (var sourceKey in sourceKeys)
                        {
                            string baseName = sourceKey.Substring(0, sourceKey.Length - sourceSuffix.Length);
                            string targetColumn = $"{baseName}{targetSuffix}";

                            if (!row.ContainsKey(targetColumn) || string.IsNullOrWhiteSpace(row[targetColumn]?.ToString()))
                            {
                                object? sourceObj = row[sourceKey];
                                string sourceText = sourceObj?.ToString() ?? string.Empty;

                                // Exceptional fallback for M.Localization.Entry: Try BaseTemplate if Template is empty
                                if (string.IsNullOrWhiteSpace(sourceText) &&
                                    sheetName.Equals("M.Localization.Entry", StringComparison.OrdinalIgnoreCase) &&
                                    baseName.Equals("M.Localization.Entry.Template", StringComparison.OrdinalIgnoreCase))
                                {
                                    string fallbackKey = $"M.Localization.Entry.BaseTemplate";
                                    if (row.TryGetValue(fallbackKey, out var fallbackObj))
                                    {
                                        sourceText = fallbackObj?.ToString() ?? string.Empty;
                                    }
                                }

                                if (!string.IsNullOrWhiteSpace(sourceText))
                                {
                                    if (shouldTranslate)
                                    {
                                        totalItems++;
                                        translationActions.Add(async () =>
                                        {
                                            var translatedText = await translationService.Translate(targetLanguage.Name, sourceText, sourceLanguage.Name);

                                            lock (row)
                                            {
                                                row[targetColumn] = translatedText;
                                            }

                                            consoleHelper.OverwriteLine($"Progress: {Interlocked.Increment(ref progress[0])}/{totalItems} | Sent to Azure: {translationService.TotalCharactersSent} chars");
                                        });
                                    }
                                    else if (shouldCopy)
                                    {
                                        copiedItems++;
                                        row[targetColumn] = sourceText;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (copiedItems > 0)
            {
                consoleHelper.Write($"Copied {copiedItems} items directly for {targetLanguage.Name}.");
            }

            if (totalItems == 0)
            {
                consoleHelper.Write($"All items for {targetLanguage.Name} are already translated or have no source text.");
                return;
            }

            consoleHelper.Write($"Translating {totalItems} items to {targetLanguage.Name}...");

            tasks.AddRange(translationActions.Select(action => ThrottledTranslateAsync(action, cancellationToken)));
            await Task.WhenAll(tasks);
            consoleHelper.Write($"\nTranslation to {targetLanguage.Name} complete.");
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