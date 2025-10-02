
using Microsoft.Graph;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Azure.Identity;
using System.Text.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

// Graph client
builder.Services.AddSingleton((sp) =>
{
    var credential = new DefaultAzureCredential();
    return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
});

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();
app.UseSwagger();
app.UseSwaggerUI();

// inbound DTO
record LinkDto(string tenantId, string roomUpn, string ewsItemId, string? restItemId, string visitorId);

// POST /link : receive mapping from add-in and normalize to Graph ids
app.MapPost("/link", async (LinkDto dto, GraphServiceClient graph) =>
{
    // 1) translateExchangeIds to get a REST immutable id for Graph
    var req = new Microsoft.Graph.Users.Item.TranslateExchangeIds.TranslateExchangeIdsPostRequestBody
    {
        InputIds = new()
        {
            new()
            {
                Id = dto.restItemId ?? dto.ewsItemId,
                SourceIdType = dto.restItemId != null ? Microsoft.Graph.Models.ExchangeIdFormat.RestId : Microsoft.Graph.Models.ExchangeIdFormat.EwsId,
                TargetIdType = Microsoft.Graph.Models.ExchangeIdFormat.RestImmutableEntryId
            }
        }
    };
    var result = await graph.Users[dto.roomUpn].TranslateExchangeIds.PostAsTranslateExchangeIdsPostResponseAsync(req);
    var graphEventId = result?.Value?.FirstOrDefault()?.TargetId ?? dto.restItemId ?? dto.ewsItemId;

    // 2) Get iCalUId (optional but recommended for cross-mailbox matching)
    var ev = await graph.Users[dto.roomUpn].Events[graphEventId].GetAsync(cfg =>
    {
        cfg.QueryParameters.Select = new[] { "id", "iCalUId", "subject", "start", "end" };
    });
    var iCalUId = ev?.ICalUId;

    // 3) TODO: Upsert into your DB (pseudo)
    Console.WriteLine($"UPSERT tenant={dto.tenantId}, room={dto.roomUpn}, eventId={graphEventId}, iCalUId={iCalUId}, visitorId={dto.visitorId}");

    // 4) (Optional) write Open Extension on the event
    // await GraphExtensions.WriteOpenExtensionAsync(graph, dto.roomUpn, graphEventId, "com.yourco.visitor", new Dictionary<string, object>{
    //     ["visitorId"] = dto.visitorId
    // });

    return Results.Accepted();
});

app.Run();
