using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Microsoft.Net.Http.Headers;
using Sitecore.ContentHub.Integration.SearchConnector.Models.Options;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Abstract;
using Sitecore.ContentHub.Integration.SearchConnector.Services.Concrete;

var builder = FunctionsApplication.CreateBuilder(args);

var queueName = builder.Configuration.GetValue<string>($"{ServiceBusOptions.ConfigurationSectionName}:{nameof(ServiceBusOptions.Name)}") ?? "";
builder.AddServiceDefaults()
    .AddAzureServiceBusClient(queueName);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddSingleton<IConfigHelper, ConfigHelper>()
    .AddScoped<IContentHubClientFactory, ContentHubClientFactory>()
    .AddTransient<IContentHubAssetService, ContentHubAssetService>()
    .AddTransient<IContentHubClientHelper, ContentHubClientHelper>()
    .AddTransient<IContentHubEntityLoadConfigurationHelper, ContentHubEntityLoadConfigurationHelper>()
    .AddTransient<IContentHubEntityService, ContentHubEntityService>()
    .AddTransient<IContentHubQueryBuilderService, ContentHubQueryBuilderService>()
    .AddTransient<IContentHubQueryService, ContentHubQueryService>()
    .AddTransient<IContentHubSearchService, ContentHubSearchService>()
    .AddTransient<ICultureHelper, CultureHelper>()
    .AddTransient<IDocumentDataService, DocumentDataService>()
    .AddTransient<IMessageWorker, MessageWorker>()
    .AddTransient<ISearchApiHelper, SearchApiHelper>();

builder.Services
    .AddOptions<ContentHubOptions>()
    .BindConfiguration(ContentHubOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<SearchOptions>()
    .BindConfiguration(SearchOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<ServiceBusOptions>()
    .BindConfiguration(ServiceBusOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddHttpClient<ISearchApiHelper, SearchApiHelper>((services, client) =>
    {
        var searchOptions = services.GetRequiredService<IOptions<SearchOptions>>().Value;

        client.BaseAddress = searchOptions.BaseUrl;
        client.DefaultRequestHeaders.Add(HeaderNames.Authorization, searchOptions.AuthToken);
    });

builder.Build().Run();
