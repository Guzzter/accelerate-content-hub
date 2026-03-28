using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.ContentHub;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract
{
    public interface IMessageWorker
    {
        public Task ProcessUpdateMessage(UpdateMessage updateMessage);
    }
}
