using Microsoft.Extensions.DependencyInjection;
using Sitecore.CH.Cli.Plugins.Base.Services.Abstract;
using Sitecore.CH.Cli.Plugins.Base.Services.Concrete;

namespace Sitecore.CH.Cli.Plugins.Base
{
    public static class Register
    {
        public static IServiceCollection AddBaseServices(this IServiceCollection services)
        {
            return services.AddTransient<IClientHelper, ClientHelper>()
                .AddTransient<IEntityHelper, EntityHelper>();
        }
    }
}
