using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Models.Options;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Abstract;
using Sitecore.ContentHub.Integration.OrderCloudConnector.Services.Concrete;

var builder = FunctionsApplication.CreateBuilder(args);

builder.AddServiceDefaults();

var busName = builder.Configuration.GetValue<string>($"{ServiceBusOptions.ConfigurationSectionName}:{nameof(ServiceBusOptions.Name)}") ?? "";
if(!string.IsNullOrWhiteSpace(busName))
    builder.AddAzureServiceBusClient(busName);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddScoped<IContentHubClientFactory, ContentHubClientFactory>()
    .AddScoped<IOrderCloudClientFactory, OrderCloudClientFactory>()
    .AddTransient<IContentHubClientHelper, ContentHubClientHelper>()
    .AddTransient<IContentHubEntityService, ContentHubEntityService>()
    .AddTransient<IOrderCloudClientHelper, OrderCloudClientHelper>()
    .AddTransient<IMessageWorker, MessageWorker>();

builder.Services
    .AddOptions<ContentHubOptions>()
    .BindConfiguration(ContentHubOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<OrderCloudOptions>()
    .BindConfiguration(OrderCloudOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<SchemaOptions>()
    .BindConfiguration(SchemaOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<ServiceBusOptions>()
    .BindConfiguration(ServiceBusOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Build().Run();
