using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sitecore.ContentHub.Integration.AssetImporter.Models.Options;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete;

var builder = FunctionsApplication.CreateBuilder(args);

builder.AddServiceDefaults();

var blobName = builder.Configuration.GetValue<string>($"{AzureOptions.ConfigurationSectionName}:{nameof(AzureOptions.StorageContainerName)}") ?? "";
builder.ConfigureFunctionsWebApplication().AddAzureBlobClient(blobName);

builder.Services
    .AddScoped<IContentHubClientFactory, ContentHubClientFactory>()
    .AddTransient<IContentHubClientHelper, ContentHubClientHelper>()
    .AddTransient<IContentHubEntityHelper, ContentHubEntityHelper>()
    .AddTransient<IContentHubSearchHelper, ContentHubSearchHelper>()
    .AddTransient<IExportWorker, ExportWorker>()
    .AddTransient<IExcelHelper, ExcelHelper>()
    .AddTransient<IStorageService, StorageService>()
    .AddTransient<IUploadService, UploadService>()
    .AddTransient<IUploadWorker, UploadWorker>();

builder.Services
    .AddOptions<AzureOptions>()
    .BindConfiguration(AzureOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Services
    .AddOptions<ContentHubOptions>()
    .BindConfiguration(ContentHubOptions.ConfigurationSectionName)
    .ValidateDataAnnotations();

builder.Build().Run();