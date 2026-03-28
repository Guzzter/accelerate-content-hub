using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Sdk.WebClient.Authentication;
using Stylelabs.M.Sdk.WebClient;
using Sitecore.ContentHub.Integration.AssetImporter.Models.Options;
using Microsoft.Extensions.Options;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ContentHubClientFactory(IOptions<ContentHubOptions> contentHubOptions) : IContentHubClientFactory
    {
        private IWebMClient? client;

        public IWebMClient CreateClient()
        {
            if (client == null)
            {
                var endpoint = contentHubOptions.Value.BaseUrl;

                var oAuth = new OAuthPasswordGrant
                {
                    ClientId = contentHubOptions.Value.ClientId,
                    ClientSecret = contentHubOptions.Value.ClientSecret,
                    UserName = contentHubOptions.Value.Username,
                    Password = contentHubOptions.Value.Password
                };

                client = MClientFactory.CreateMClient(endpoint, oAuth);
            }
            return client;
        }
    }
}
