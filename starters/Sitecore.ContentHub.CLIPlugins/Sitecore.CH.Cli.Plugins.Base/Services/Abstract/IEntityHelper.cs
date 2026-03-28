using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Contracts.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Sitecore.CH.Cli.Plugins.Base.Services.Abstract
{
    public interface IEntityHelper
    {
        Task<IEntity> Get(long id, IEntityLoadConfiguration? loadConfig = null);

        Task<IEntity> Get(string identifier, IEntityLoadConfiguration? loadConfig = null);
    }
}
