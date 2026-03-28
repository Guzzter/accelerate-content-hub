using Microsoft.Extensions.Logging;
using Sitecore.ContentHub.Integration.AssetImporter.Constants;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Sdk.Exceptions;
using Stylelabs.M.Sdk.WebClient;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ContentHubClientHelper(ILoggerFactory loggerFactory, IContentHubClientFactory clientFactory) : IContentHubClientHelper
    {
        private readonly ILogger logger = loggerFactory.CreateLogger<ContentHubClientHelper>();
        private readonly IWebMClient client = clientFactory.CreateClient();
        private readonly Random random = new();

        public async Task Execute(Func<IWebMClient, Task> delegateFunc)
        {
            await RetryWrapper(c => { delegateFunc(c); return Task.FromResult(1); }, 1);
        }

        public async Task<T> Execute<T>(Func<IWebMClient, Task<T>> delegateFunc)
        {
            return await RetryWrapper(delegateFunc, 1);
        }

        public T ExecuteSync<T>(Func<IWebMClient, T> delegateFunc)
        {
            var task = Task.Run(() => RetryWrapper(c => Task.FromResult(delegateFunc(c)), 1));
            if (task.IsFaulted && task.Exception != null)
                throw task.Exception;
            return task.Result;
        }

        private async Task<T> RetryWrapper<T>(Func<IWebMClient, Task<T>> delegateFunc, int attemptNumber)
        {
            try
            {
                return await delegateFunc(client);
            }
            catch (WebApiException exception)
            {
                var delay = ShouldRetry(exception, attemptNumber);
                if (delay.HasValue)
                {
                    await Task.Delay(delay.Value);
                    return await RetryWrapper(delegateFunc, attemptNumber++);
                }
                else
                    throw;
            }
        }

        private int? ShouldRetry(WebApiException exception, int attemptNumber)
        {
            if (exception.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
            {
                if (attemptNumber < PlatformConstants.Client.RequestRetryAttempts)
                {
                    var delay = random.Next(PlatformConstants.Client.RequestRetryDelayMin, PlatformConstants.Client.RequestRetryDelayMax);
                    logger.LogWarning($"Too Many Requests exception received on attempt {attemptNumber}/{PlatformConstants.Client.RequestRetryAttempts} delaying for {delay}ms");
                    return delay;
                }
                logger.LogError($"Too Many Requests exception received on attempt {attemptNumber}/{PlatformConstants.Client.RequestRetryAttempts}. Exiting.");
            }
            return null;
        }
    }
}
