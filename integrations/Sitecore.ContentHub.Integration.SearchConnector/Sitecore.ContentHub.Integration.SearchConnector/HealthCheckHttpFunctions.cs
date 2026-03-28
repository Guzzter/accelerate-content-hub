using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace Sitecore.ContentHub.Integration.SearchConnector
{
    public class HealthCheckHttpFunctions(ILogger<HealthCheckHttpFunctions> logger)
    {
        [Function("HealthCheck")]
        public IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
        {
            logger.LogInformation("Health check requested");
            return new OkObjectResult("OK");
        }
    }
}
