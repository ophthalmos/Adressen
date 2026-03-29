//using System.Net.Http.Headers;

//namespace Adressen.cls;

//internal static class HttpService
//{
//    public static readonly HttpClient Client; // Die einzige Instanz von HttpClient für die gesamte Anwendung.

//    static HttpService()
//    {
//        Client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) }; // Timeout (Standard ist 100 Sekunden)
//        Client.DefaultRequestHeaders.Accept.Clear();
//        Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
//        Client.DefaultRequestHeaders.UserAgent.ParseAdd("Adressen/1.0");
//    }
//}
using System.Net.Http.Headers;

namespace Adressen.cls;

internal static class HttpService
{
    // Die einzige Instanz für die gesamte App-Laufzeit
    public static readonly HttpClient Client;

    static HttpService()
    {
        // Der SocketsHttpHandler ist der Motor unter .NET 10
        var handler = new SocketsHttpHandler
        {
            // Erneuert Verbindungen regelmäßig (gut für DNS-Rotation)
            PooledConnectionLifetime = TimeSpan.FromMinutes(2),

            // Wie lange darf der reine Verbindungsaufbau dauern?
            ConnectTimeout = TimeSpan.FromSeconds(5),

            // Maximale parallele Verbindungen zum selben Host (z.B. Google)
            MaxConnectionsPerServer = 20,

            // Erlaubt das schnellere HTTP/2 (und falls unterstützt HTTP/3)
            EnableMultipleHttp2Connections = true
        };

        Client = new HttpClient(handler)
        {
            // Total-Timeout für den gesamten Request
            Timeout = TimeSpan.FromSeconds(15)
        };

        Client.DefaultRequestHeaders.Accept.Clear();
        Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        Client.DefaultRequestHeaders.UserAgent.ParseAdd("Adressen/1.0");
    }
}