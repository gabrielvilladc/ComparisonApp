using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using ComparisonApp;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Collections.Generic;
using ComparisonApp.Entities;
using ComparisonApp.Entities.Performance;
using ComparisonApp.Entities.PortfolioStats;
using ComparisonApp.Functions;

var builder = new ConfigurationBuilder();
BuildConfig(builder);

IConfiguration configuration = builder.Build();

using IHost host = Host.CreateDefaultBuilder(args)
            .ConfigureServices((_, services) =>
                services
                //.AddTransient<ITransientOperation, DefaultOperation>()
                //.AddScoped<IScopedOperation, DefaultOperation>()
                .AddAutoMapper(typeof(Program).Assembly)
                .AddSingleton<ISingletonOperation, DefaultOperation>()
                .AddTransient<OperationLogger>()
                .AddTransient<HoldingsLogic>()
                .AddTransient<PerformanceLogic>()
                .AddTransient<PortfolioStatLogic>()
                .AddTransient<GeneralLogic>())
                .Build();

using IServiceScope serviceScope = host.Services.CreateScope();
IServiceProvider provider = serviceScope.ServiceProvider;

//DI
HoldingsLogic holdings = provider.GetRequiredService<HoldingsLogic>();
PerformanceLogic performance = provider.GetRequiredService<PerformanceLogic>();
PortfolioStatLogic portfolioStats = provider.GetRequiredService<PortfolioStatLogic>();
GeneralLogic generalLogic = provider.GetRequiredService<GeneralLogic>();

//Generating all the reports list with the differences
List<FinalHoldingsComparison> _listFinalHoldingsDifferences = holdings.GetHoldingsDifferences(host.Services, args);
List<FinalPerformanceComparison> _listFinalPerformanceDifferences = performance.GetPerformanceDifferences(host.Services, args);
List<FinalPortfolioStatsComparison> _listFinalPortfolioStatsDifferences = portfolioStats.GetPortfoliostatsDifferences(host.Services, args);

//Generating file with differences
generalLogic.GenerateReport(_listFinalHoldingsDifferences, _listFinalPerformanceDifferences, _listFinalPortfolioStatsDifferences, args[2], args[1]);
//await host.RunAsync();

static void BuildConfig(IConfigurationBuilder builder)
{
    builder.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "Production"}.json", optional: true)
        .AddEnvironmentVariables()
        .Build();
}

