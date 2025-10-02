
using Microsoft.Graph;

public static class GraphExtensions
{
    public static async Task WriteOpenExtensionAsync(GraphServiceClient graph, string userOrRoom, string eventId, string extName, IDictionary<string, object> payload)
    {
        var body = new Microsoft.Graph.Models.OpenTypeExtension
        {
            OdataType = "#microsoft.graph.openTypeExtension",
            ExtensionName = extName,
            AdditionalData = payload
        };

        await graph.Users[userOrRoom].Events[eventId].Extensions.PostAsync(body);
    }
}
