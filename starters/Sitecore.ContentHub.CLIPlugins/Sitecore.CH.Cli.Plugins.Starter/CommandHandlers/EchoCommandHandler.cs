using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Core.Abstractions.Rendering;
using Sitecore.CH.Cli.Plugins.Starter.Model.Parameters;
using Stylelabs.M.Sdk.WebClient;
using System.CommandLine.Invocation;

namespace Sitecore.CH.Cli.Plugins.Starter.CommandHandlers
{
    internal class EchoCommandHandler(ILogger<HelloCommandHandler> logger, Lazy<IWebMClient> client, IOutputRenderer renderer, IOptions<EchoParameters> parameters) : BaseCommandHandler(client, renderer)
    {
        private readonly ILogger<HelloCommandHandler> logger = logger;
        private readonly EchoParameters parameters = parameters.Value;

        public async override Task<int> InvokeAsync(InvocationContext context)
        {
            Renderer.WriteLine($"Your message was: {parameters.Message}");
            await Task.Delay(5000);
            return 1;
        }
    }
}
