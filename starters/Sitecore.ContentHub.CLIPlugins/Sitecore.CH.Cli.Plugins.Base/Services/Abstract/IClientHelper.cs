using Stylelabs.M.Sdk.WebClient;

namespace Sitecore.CH.Cli.Plugins.Base.Services.Abstract
{
    public interface IClientHelper
    {
        Task Execute(Func<IWebMClient, Task> delegateFunc);

        Task<T> Execute<T>(Func<IWebMClient, Task<T>> delegateFunc);

        T ExecuteSync<T>(Func<IWebMClient, T> delegateFunc);
    }
}
