
using Azure.Identity;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        services.AddSingleton(sp => new GraphServiceClient(new DefaultAzureCredential(), new[] { "https://graph.microsoft.com/.default" }));
    })
    .Build();

host.Run();
