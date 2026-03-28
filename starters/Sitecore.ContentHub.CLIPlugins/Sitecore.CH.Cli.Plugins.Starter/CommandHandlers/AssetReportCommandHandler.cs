using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Core.Abstractions.Rendering;
using Sitecore.CH.Cli.Plugins.Starter.Model;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Framework.Essentials.LoadOptions;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.WebClient;
using System.Collections.Concurrent;
using System.CommandLine.Invocation;
using System.Dynamic;
using System.Globalization;
using System.Text;

namespace Sitecore.CH.Cli.Plugins.Starter.CommandHandlers
{
    internal class AssetReportCommandHandler(
        ILogger<AssetReportCommandHandler> logger,
        Lazy<IWebMClient> client,
        IOutputRenderer renderer) : BaseCommandHandler(client, renderer)
    {
        private readonly ILogger logger = logger;
        private readonly ConcurrentBag<AssetReportEntry> assetsCsvReport = new();

        public override async Task<int> InvokeAsync(InvocationContext context)
        {
            try
            {
                logger.LogInformation("Starting asset report generation");

                await Renderer.ShowLoadingIndicatorAsync(
                    "Loading and processing assets",
                    async () =>
                    {
                        await LoadAssetsAsync();
                        return true;
                    });

                logger.LogInformation($"Successfully processed {assetsCsvReport.Count} assets");

                // Output the report to CSV file
                if (assetsCsvReport.Any())
                {
                    var outputPath = await WriteCsvReportAsync(assetsCsvReport.ToList());
                    logger.LogInformation($"Report generated with {assetsCsvReport.Count} entries");
                    logger.LogInformation($"CSV file written to: {outputPath}");
                    Renderer.WriteLine($"Report saved to: {outputPath}");
                }
                else
                {
                    logger.LogWarning("No assets found to report on");
                }

                return 1;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An error occurred while generating asset report");
                return 0;
            }
        }

        private async Task LoadAssetsAsync()
        {
            const int take = 100;
            
            var mClient = Client.Value;

            var query = Query.CreateQuery(entities => from e in entities
                where e.DefinitionName == "M.Asset"
                orderby e.CreatedOn ascending
                select e);
            query.Take = take;

            var assetLoadConfig = new EntityLoadConfiguration(
                CultureLoadOption.Default,
                new PropertyLoadOption("AssetId", "FileName", "Renditions"),
                new RelationLoadOption("FinalLifeCycleStatusToAsset", "AssetTypeToAsset", "AssetMediaToAsset"));

            var queryResult = await mClient.Querying.QueryAsync(query, assetLoadConfig).ConfigureAwait(false);

            while (queryResult.Items.Any())
            {
                foreach (var asset in queryResult.Items)
                {
                    if (!asset.Id.HasValue || assetsCsvReport.Any(x => x.EntityId == asset.Id.Value))
                        continue;

                    ProcessAsset(asset);
                }

                // Load next page
                if (queryResult.Items.Count < take)
                    break;

                query.Skip = (query.Skip ?? 0) + take;
                queryResult = await mClient.Querying.QueryAsync(query, assetLoadConfig).ConfigureAwait(false);
            }
        }

        private void ProcessAsset(IEntity asset)
        {
            if (!asset.Id.HasValue) return;

            var lifeCycleId = asset.GetRelation("FinalLifeCycleStatusToAsset")?.GetIds();
            var assetTypeId = asset.GetRelation("AssetTypeToAsset")?.GetIds();
            var renditionsProp = asset.GetPropertyValue<JObject>("Renditions");

            var entry = new AssetReportEntry
            {
                EntityId = asset.Id.Value,
                Identifier = asset.Identifier,
                FileName = asset.GetPropertyValue<string>("FileName"),
                FinalLifecycle = GetLifecycleStatus(lifeCycleId?.FirstOrDefault() ?? 0)
            };

            // Asset Media
            var assetMedia = asset.GetRelation("AssetMediaToAsset")?.GetIds();
            if (assetMedia != null && assetMedia.Any())
                entry.AssetMedia = true;

            // Fetch - original rendition present
            if (asset.Renditions != null && asset.Renditions.Any(x => x.Name.Equals("downloadOriginal")))
                entry.FetchDone = true;

            // Renditions triggered and completed
            if (asset.Renditions != null && asset.Renditions.Any())
            {
                entry.RenditionsTriggered = true;
                entry.AllRenditionsCompleted = true;
            }

            // Preview rendition present
            if (asset.Renditions != null && asset.Renditions.Any(x => x.Name.Equals("preview")))
                entry.PreviewRenditionDone = true;

            // Check for incomplete renditions
            if (renditionsProp != null && !string.IsNullOrEmpty(renditionsProp.ToString()))
            {
                var converter = new ExpandoObjectConverter();
                dynamic? renditionsJson =
                    JsonConvert.DeserializeObject<ExpandoObject>(renditionsProp.ToString(), converter);

                var foundBrokenRendition = false;
                
                if (renditionsJson != null)
                {
                    foreach (KeyValuePair<string, object> renditionsKvp in renditionsJson)
                    {
                        if (renditionsKvp.Value is not ExpandoObject value) continue;
                        
                        foreach (var rendition in value)
                        {
                            if (string.IsNullOrEmpty(rendition.Key) || !rendition.Key.Equals("status")) continue;
                            
                            if (rendition.Value == null) continue;

                            var renditionStatus = rendition.Value.ToString();
                            
                            if (string.IsNullOrEmpty(renditionStatus) || renditionStatus.Equals("completed")) continue;
                            
                            foundBrokenRendition = true;
                            entry.AllRenditionsCompleted = false;
                        }

                        if (foundBrokenRendition) break;
                    }
                }
            }

            // MetadataDone = Asset Type is set
            if (assetTypeId != null && assetTypeId.Any())
                entry.MetadataDone = true;
            
            assetsCsvReport.Add(entry);
        }

        private static string GetLifecycleStatus(long id)
        {
            return id switch
            {
                542 => "Created",
                543 => "Under Review",
                544 => "Approved",
                545 => "Rejected",
                546 => "Archived",
                _ => string.Empty
            };
        }

        private static async Task<string> WriteCsvReportAsync(List<AssetReportEntry> entries)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
            var outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "ch-cli-reports");
            Directory.CreateDirectory(outputDir);

            var outputPath = Path.Combine(outputDir, $"asset-report_{timestamp}.csv");

            var csv = new StringBuilder();

            // Write header
            csv.AppendLine(
                "Entity ID,Identifier,File Name,Status,Asset Media,Fetched,Renditions Triggered,Preview Rendition,All Renditions Completed,Metadata");

            // Write data rows
            foreach (var entry in entries)
            {
                csv.AppendLine($"{entry.EntityId}," +
                               $"{EscapeCsvField(entry.Identifier)}," +
                               $"{EscapeCsvField(entry.FileName)}," +
                               $"{EscapeCsvField(entry.FinalLifecycle)}," +
                               $"{entry.AssetMedia}," +
                               $"{entry.FetchDone}," +
                               $"{entry.RenditionsTriggered}," +
                               $"{entry.PreviewRenditionDone}," +
                               $"{entry.AllRenditionsCompleted}," +
                               $"{entry.MetadataDone}");
            }

            await File.WriteAllTextAsync(outputPath, csv.ToString()).ConfigureAwait(false);
            return outputPath;
        }

        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return string.Empty;

            // If field contains comma, quote, or newline, wrap in quotes and escape internal quotes
            if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }

            return field;
        }
    }
}