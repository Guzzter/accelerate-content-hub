var builder = DistributedApplication.CreateBuilder(args);

var serviceBus = builder.AddAzureServiceBus("messaging").RunAsEmulator();
var updateQueue = serviceBus.AddServiceBusQueue("update");

builder.AddAzureFunctionsProject<Projects.Sitecore_ContentHub_Integration_OrderCloudConnector>("functions")
    .WithReference(serviceBus)
    .WaitFor(serviceBus)
    .WithEnvironment("ServiceBus:Name", serviceBus.Resource.Name)
    .WithEnvironment("ServiceBus:UpdateQueueName", updateQueue.Resource.QueueName);

builder.Build().Run();
