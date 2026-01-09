using System.Net.Http.Headers;

namespace Adressen.cls;

internal static class HttpService
{
    public static readonly HttpClient Client; // Die einzige Instanz von HttpClient für die gesamte Anwendung.

    static HttpService()
    {
        Client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) }; // Timeout (Standard ist 100 Sekunden)
        Client.DefaultRequestHeaders.Accept.Clear();
        Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        Client.DefaultRequestHeaders.UserAgent.ParseAdd("Adressen/1.0");
    }
}
