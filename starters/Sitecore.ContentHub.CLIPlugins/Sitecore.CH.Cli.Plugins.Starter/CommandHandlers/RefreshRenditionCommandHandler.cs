using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Linq;
using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Core.Abstractions.Rendering;
using Sitecore.CH.Cli.Plugins.Base.Services.Abstract;
using Sitecore.CH.Cli.Plugins.Starter.Model.Parameters;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using Stylelabs.M.Sdk.WebClient;
using System.CommandLine.Invocation;

namespace Sitecore.CH.Cli.Plugins.Starter.CommandHandlers
{
    internal class RefreshRenditionCommandHandler(
        ILogger<RefreshRenditionCommandHandler> logger,
        Lazy<IWebMClient> client,
        IOutputRenderer renderer,
        IOptions<GetEntityParameters> parameters,
        IEntityHelper entityHelper) : BaseCommandHandler(client, renderer)
    {
        private readonly ILogger logger = logger;
        private readonly GetEntityParameters parameters = parameters.Value;
        private readonly IEntityHelper entityHelper = entityHelper;

        public override async Task<int> InvokeAsync(InvocationContext context)
        {
            // Validate parameters
            var paramCount = (parameters.Id.HasValue ? 1 : 0) + 
                           (!string.IsNullOrEmpty(parameters.Identifier) ? 1 : 0) + 
                           (!string.IsNullOrEmpty(parameters.File) ? 1 : 0);
            
            if (paramCount == 0)
                throw new ArgumentException("Either Id, Identifier, or File must be provided as arguments");
            if (paramCount > 1)
                throw new ArgumentException("Only one of Id, Identifier, or File must be provided as arguments");

            try
            {
                List<long> targetIds = new List<long>();

                // Load IDs based on the parameter type
                if (!string.IsNullOrEmpty(parameters.File))
                {
                    targetIds = await Renderer.ShowLoadingIndicatorAsync($"Loading entity IDs from file: {parameters.File}",
                        () => Task.FromResult(LoadIdsFromFile(parameters.File)));
                    
                    if (targetIds.Count == 0)
                    {
                        logger.LogWarning("No valid IDs found in the file");
                        return 0;
                    }
                    
                    logger.LogInformation($"Loaded {targetIds.Count} entity IDs from file");
                }
                else
                {
                    var entity = await Renderer.ShowLoadingIndicatorAsync("Getting entity",
                        () => parameters.Id.HasValue
                            ? entityHelper.Get(parameters.Id.Value)
                            : entityHelper.Get(parameters.Identifier));

                    if (!entity.Id.HasValue)
                    {
                        throw new InvalidOperationException("Entity does not have a valid ID");
                    }
                    
                    targetIds.Add(entity.Id.Value);
                }

                // Create refresh renditions job
                var jobId = await Renderer.ShowLoadingIndicatorAsync(
                    targetIds.Count == 1 
                        ? $"Creating refresh renditions job for entity {targetIds[0]}"
                        : $"Creating refresh renditions job for {targetIds.Count} entities",
                    () => CreateRefreshRenditionsJobAsync(targetIds));

                logger.LogInformation($"Successfully created refresh renditions job {jobId} for {targetIds.Count} entity(ies)");

                return 1;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An error occurred");
                return 0;
            }
        }

        private List<long> LoadIdsFromFile(string filePath)
        {
            var ids = new List<long>();
            
            if (!System.IO.File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }

            string[] lines = System.IO.File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                if (long.TryParse(line.Trim(), out long assetId))
                {
                    ids.Add(assetId);
                }
                else if (!string.IsNullOrWhiteSpace(line.Trim()))
                {
                    logger.LogWarning($"Skipping invalid ID: {line.Trim()}");
                }
            }

            return ids;
        }

        private async Task<long> CreateRefreshRenditionsJobAsync(List<long> targetIds)
        {
            try
            {
                var mClient = Client.Value;

                // First create the job entity
                var jobEntity = await mClient.EntityFactory.CreateAsync("M.Job").ConfigureAwait(false);
                jobEntity.SetPropertyValue("Job.Type", "MassEdit");
                jobEntity.SetPropertyValue("Job.State", "Created");
                jobEntity.SetPropertyValue("Job.Condition", "Pending");
                jobEntity.SetPropertyValue("Job.TargetCount", Convert.ToInt64(0));
                jobEntity.SetPropertyValue("Job.TargetsCompleted", Convert.ToInt64(0));

                var jobId = await mClient.Entities.SaveAsync(jobEntity).ConfigureAwait(false);
                logger.LogInformation($"Created job entity with ID: {jobId}");

                // Then create the job description
                var descriptionEntity = await mClient.EntityFactory.CreateAsync("M.JobDescription").ConfigureAwait(false);

                string jobConfiguration = $@"{{
                ""$type"": ""Stylelabs.M.Base.MassEdit.MassEditJobDescription, Stylelabs.M.Base"",
                ""Operations"": [
                    {{
                        ""$type"": ""Stylelabs.M.Base.MassEdit.RefreshRenditionsOperation, Stylelabs.M.Base"",
                        ""Renditions"": [],
				        ""FailedOnly"": false,
				        ""RefreshHistory"": false,
				        ""RefreshSubfiles"": false
			        }}
                ],
                ""FinalizeOperations"": [],
                ""Targets"": [{string.Join(",", targetIds)}]
            }}";
                
                descriptionEntity.SetPropertyValue("Job.Configuration", JToken.Parse(jobConfiguration));
                var jobRelation = descriptionEntity.GetRelation<IChildToOneParentRelation>("JobToJobDescription");
                jobRelation.SetId(jobId);

                await mClient.Entities.SaveAsync(descriptionEntity).ConfigureAwait(false);
                logger.LogInformation("Created job description entity");

                // Finally update the job to trigger execution
                var loadConfig = EntityLoadConfiguration.Full;

                // Wait a bit for the job description to be fully processed
                await Task.Delay(TimeSpan.FromSeconds(8)).ConfigureAwait(false);
                
                jobEntity = await mClient.Entities.GetAsync(jobId, loadConfig).ConfigureAwait(false);
                jobEntity.SetPropertyValue("Job.State", "Pending");
                jobEntity.SetPropertyValue("Job.TargetCount", Convert.ToInt64(targetIds.Count));
                await mClient.Entities.SaveAsync(jobEntity).ConfigureAwait(false);
                
                logger.LogInformation($"Job {jobId} updated to Pending state");
                
                return jobId;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error creating refresh renditions job");
                throw;
            }
        }
    }
}