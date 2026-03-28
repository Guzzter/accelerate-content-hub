using Microsoft.Extensions.DependencyInjection;
using Sitecore.CH.Cli.Core.Abstractions.Infrastructure;
using Sitecore.CH.Cli.Core.Extensions;
using Sitecore.CH.Cli.Plugins.Base;
using Sitecore.CH.Cli.Plugins.Starter.CommandHandlers;
using Sitecore.CH.Cli.Plugins.Starter.Commands;
using Sitecore.CH.Cli.Plugins.Starter.Model.Parameters;

namespace Sitecore.CH.Cli.Plugins.Starter
{
    public class Plugin : IPlugin
    {
        public void ConfigureServices(IServiceCollection services)
        {
            services
                .AddCommandHandler<HelloCommandHandler>()
                .AddCommandHandler<EchoCommandHandler, EchoParameters>()
                .AddCommandHandler<GetEntityCommandHandler, GetEntityParameters>()
                .AddCommandHandler<RefreshRenditionCommandHandler, GetEntityParameters>()
                .AddCommandHandler<AssetReportCommandHandler>()
                .AddBaseServices();
        }

        public void RegisterCommands(ICommandRegistry registry)
        {
            registry.RegisterCommandGroup("starter",
                [
                    new HelloCommand(),
                    new EchoCommand(),
                    new GetEntityCommand(),
                    new RefreshRenditionCommand(),
                    new AssetReportCommand(),
                ],
                "CLI plugin starter test commands.");
        }
    }
}