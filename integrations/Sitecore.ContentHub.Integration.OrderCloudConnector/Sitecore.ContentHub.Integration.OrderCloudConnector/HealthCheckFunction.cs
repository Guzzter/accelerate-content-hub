using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace Sitecore.ContentHub.Integration.OrderCloudConnector;

public class HealthCheckFunction(ILogger<HealthCheckFunction> logger)
{
    [Function("HealthCheck")]
    public IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequest req)
    {
        logger.LogInformation("Health check requested");
        return new OkObjectResult("OK");
    }
}