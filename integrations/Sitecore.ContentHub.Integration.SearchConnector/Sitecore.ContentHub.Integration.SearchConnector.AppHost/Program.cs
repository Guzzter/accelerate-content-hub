var builder = DistributedApplication.CreateBuilder(args);

var serviceBus = builder.AddAzureServiceBus("messaging").RunAsEmulator();
var upsertQueue = serviceBus.AddServiceBusQueue("upsert");
var deleteQueue = serviceBus.AddServiceBusQueue("delete");

builder.AddAzureFunctionsProject<Projects.Sitecore_ContentHub_Integration_SearchConnector>("functions")
        .WithReference(serviceBus)
        .WaitFor(serviceBus)
        .WithEnvironment("ServiceBus:Name", serviceBus.Resource.Name)
        .WithEnvironment("ServiceBus:UpsertQueueName", upsertQueue.Resource.QueueName)
        .WithEnvironment("ServiceBus:DeleteQueueName", deleteQueue.Resource.QueueName);

builder.Build().Run();
