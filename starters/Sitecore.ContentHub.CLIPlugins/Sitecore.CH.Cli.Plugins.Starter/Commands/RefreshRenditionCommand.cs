using Sitecore.CH.Cli.Core.Abstractions.Commands;
using Sitecore.CH.Cli.Plugins.Starter.CommandHandlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sitecore.CH.Cli.Plugins.Starter.Commands
{
    internal class RefreshRenditionCommand : BaseCommand<RefreshRenditionCommandHandler>
    {
        public RefreshRenditionCommand() : base("refresh-rendition", "A command to refresh renditions on an entity by its id or identifier")
        {
            AddOption<long>("The ID of the entity to refresh", false, "--id", "-i");
            AddOption<string>("The identifier of the entity to refresh", false, "--identifier");
            AddOption<string>("Path to a text file containing entity IDs (one per line)", false, "--file", "-f");
        }
    }
}
