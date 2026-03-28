using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Plugins.Starter.CommandHandlers;

namespace Sitecore.CH.Cli.Plugins.Starter.Commands
{
    internal class AssetReportCommand()
        : BaseCommand<AssetReportCommandHandler>("report-all-assets", "Create report of state of Assets in the system")
    {
        // No additional options needed for this report command
    }
}
