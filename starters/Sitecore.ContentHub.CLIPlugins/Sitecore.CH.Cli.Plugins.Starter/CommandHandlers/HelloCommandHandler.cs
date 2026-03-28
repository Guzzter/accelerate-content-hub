using Microsoft.Extensions.Logging;
using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Core.Abstractions.Rendering;
using Stylelabs.M.Sdk.WebClient;
using System.CommandLine.Invocation;

namespace Sitecore.CH.Cli.Plugins.Starter.CommandHandlers
{
    internal class HelloCommandHandler : BaseCommandHandler
    {
        private readonly ILogger<HelloCommandHandler> logger;

        public HelloCommandHandler(ILogger<HelloCommandHandler> logger, Lazy<IWebMClient> client, IOutputRenderer renderer) : base(client, renderer)
        {
            this.logger = logger;
        }

        public async override Task<int> InvokeAsync(InvocationContext context)
        {
            Renderer.WriteLine("Hello World :-)");
            await Task.Delay(5000);
            return 1;
        }
    }
}
