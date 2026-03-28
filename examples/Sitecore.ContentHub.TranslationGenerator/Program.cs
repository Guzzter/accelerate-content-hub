using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sitecore.CH.TranslationGenerator.Models;
using Sitecore.CH.TranslationGenerator.Services.Abstract;
using Sitecore.CH.TranslationGenerator.Services.Concrete;

var builder = Host.CreateApplicationBuilder();

builder.Services
    .AddHostedService<ApplicationWorker>()
    .AddTransient<IConsoleHelper, ConsoleHelper>()
    .AddTransient<IExcelService, ExcelService>()
    .AddTransient<ITranslationService, TranslationService>()

    //.AddTransient<ITranslationService, MockTranslationService>();
    .AddOptions<TranslationSettings>().BindConfiguration(TranslationSettings.Key);

builder.Configuration
    .AddJsonFile("appsettings.local.json", true, true);

builder.Build().Run();