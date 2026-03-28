using Microsoft.Extensions.Logging;
using Sitecore.CH.Cli.Plugins.Base.Services.Abstract;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;

namespace Sitecore.CH.Cli.Plugins.Base.Services.Concrete
{
    internal class EntityHelper : IEntityHelper
    {
        private readonly ILogger logger;
        private readonly IClientHelper clientHelper;

        public EntityHelper(ILogger<EntityHelper> logger, IClientHelper clientHelper)
        {
            this.logger = logger;
            this.clientHelper = clientHelper;
        }

        public async Task<IEntity> Get(long id, IEntityLoadConfiguration? loadConfig = null)
        {
            return await clientHelper.Execute(client => client.Entities.GetAsync(id, loadConfig));
        }

        public async Task<IEntity> Get(string identifier, IEntityLoadConfiguration? loadConfig = null)
        {
            return await clientHelper.Execute(client => client.Entities.GetAsync(identifier, loadConfig));
        }
    }
}
