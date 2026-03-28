using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Core.Abstractions.Rendering;
using Sitecore.CH.Cli.Plugins.Base.Services.Abstract;
using Sitecore.CH.Cli.Plugins.Starter.Model.Parameters;
using Stylelabs.M.Sdk.WebClient;
using System.CommandLine.Invocation;

namespace Sitecore.CH.Cli.Plugins.Starter.CommandHandlers
{
    internal class GetEntityCommandHandler(ILogger<GetEntityCommandHandler> logger, Lazy<IWebMClient> client, IOutputRenderer renderer, IOptions<GetEntityParameters> parameters, IEntityHelper entityHelper) : BaseCommandHandler(client, renderer)
    {
        private readonly ILogger logger = logger;
        private readonly GetEntityParameters parameters = parameters.Value;
        private readonly IEntityHelper entityHelper = entityHelper;

        public async override Task<int> InvokeAsync(InvocationContext context)
        {
            if (!parameters.Id.HasValue && string.IsNullOrEmpty(parameters.Identifier))
                throw new ArgumentException("Either Id or Identifier must be provided as arguments");
            if (parameters.Id.HasValue && !string.IsNullOrEmpty(parameters.Identifier))
                throw new ArgumentException("Only one of Id or Identifier must be provided as arguments");

            try
            {
                var entity = await Renderer.ShowLoadingIndicatorAsync("Getting entity", () => parameters.Id.HasValue ? entityHelper.Get(parameters.Id.Value) : entityHelper.Get(parameters.Identifier));
                Renderer.RenderJson(entity);
                await Task.Delay(5000);
                return 1;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An error occurred");
                return 0;
            }
        }
    }
}
