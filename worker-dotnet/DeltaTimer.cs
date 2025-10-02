
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

public class DeltaTimer
{
    private readonly GraphServiceClient _graph;
    private readonly ILogger _logger;

    public DeltaTimer(GraphServiceClient graph, ILoggerFactory lf)
    {
        _graph = graph;
        _logger = lf.CreateLogger<DeltaTimer>();
    }

    // Every 5 minutes
    [Function("DeltaTimer")]
    public async Task Run([TimerTrigger("0 */5 * * * *")] TimerInfo timer)
    {
        var rooms = Environment.GetEnvironmentVariable("ROOM_UPNS")?.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries) ?? Array.Empty<string>();
        var daysPast = int.TryParse(Environment.GetEnvironmentVariable("WINDOW_DAYS_PAST"), out var p) ? p : 1;
        var daysFuture = int.TryParse(Environment.GetEnvironmentVariable("WINDOW_DAYS_FUTURE"), out var f) ? f : 7;

        foreach (var room in rooms)
        {
            try
            {
                var start = DateTimeOffset.UtcNow.Date.AddDays(-daysPast).ToString("o");
                var end = DateTimeOffset.UtcNow.Date.AddDays(daysFuture).ToString("o");

                var page = await _graph.Users[room].CalendarView.Delta.GetAsync(cfg =>
                {
                    cfg.QueryParameters.StartDateTime = start;
                    cfg.QueryParameters.EndDateTime = end;
                    cfg.Headers.Add("Prefer", "outlook.timezone=\"Tokyo Standard Time\"");
                });

                while (page is not null)
                {
                    foreach (var ev in page.Value ?? new List<Microsoft.Graph.Models.Event>())
                    {
                        // TODO: JOIN your DB by ev.Id or ev.ICalUId, then upsert cache
                        _logger.LogInformation("Room {room}: {subject} ({start} - {end})", room, ev.Subject, ev.Start?.DateTime, ev.End?.DateTime);
                    }

                    if (!string.IsNullOrEmpty(page.OdataNextLink))
                        page = await _graph.Users[room].CalendarView.Delta.WithUrl(page.OdataNextLink).GetAsync();
                    else
                        break;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Delta failed for room {room}", room);
            }
        }
    }
}
