using System.Net.Http.Headers;

namespace Adressen.cls;

internal static class HttpService
{
    public static readonly HttpClient Client;  // Die einzige Instanz für die gesamte App-Laufzeit

    static HttpService()
    {
        var handler = new SocketsHttpHandler  // Der SocketsHttpHandler ist der Motor unter .NET 10
        {
            PooledConnectionLifetime = TimeSpan.FromMinutes(2),  // Erneuert Verbindungen regelmäßig (gut für DNS-Rotation)
            ConnectTimeout = TimeSpan.FromSeconds(5),  // Wie lange darf der reine Verbindungsaufbau dauern?
            MaxConnectionsPerServer = 20,  // Maximale parallele Verbindungen zum selben Host (z.B. Google)
            EnableMultipleHttp2Connections = true  // Erlaubt das schnellere HTTP/2 (und falls unterstützt HTTP/3)
        };
        Client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(15) };  // Total-Timeout für den gesamten Request
        Client.DefaultRequestHeaders.Accept.Clear();
        Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        Client.DefaultRequestHeaders.UserAgent.ParseAdd("Adressen/1.0");
    }
}